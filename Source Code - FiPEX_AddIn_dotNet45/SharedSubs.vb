Imports System.Drawing
Imports ESRI.ArcGIS.Geodatabase
Imports ESRI.ArcGIS.Carto
Imports ESRI.ArcGIS.Framework
Imports ESRI.ArcGIS.ArcMapUI
Imports ESRI.ArcGIS.esriSystem
Imports ESRI.ArcGIS.EditorExt
Imports ESRI.ArcGIS.Geometry
Imports ESRI.ArcGIS.NetworkAnalysis
Imports ESRI.ArcGIS.Display

Public Class SharedSubs

    Public Shared Sub BreadthFirstSearch(ByRef pOriginalBarriersList As IEnumNetEID,
                                       ByRef pOriginaljuncFlagsList As IEnumNetEID,
                                       ByRef m_UtilityNetworkAnalysisExt As ESRI.ArcGIS.EditorExt.IUtilityNetworkAnalysisExt,
                                       ByRef sType As String, ByRef sFieldName As String, ByRef bContinue As Boolean)
        ' GO - Sep 2021
        ' ######################### Depth first search ##############
        ' Uses ArcMap UNA traces to perform DFS upstream from sink(s)
        ' - Labels by saving selection-set object then edit session
        ' - Based on loops used in the analysis.vb function 

        ' sType -- can be 'segment', 'subsegment' - determines analysis types
        '   SEGMENT - upstream analysis, edges labelled from upstream trace of user-set nodes only (no branches, no sources)
        '   SUBSEGMENT - downstream analysis from all sources, branches, barriers - edges labelled, flags / barrier / branch nodes labelled
        ' sFieldName -- is the fieldname to insert

        'Dim sHGSegmentFieldName As String = "FIPEX_HG_SegID"
        Dim sValueType As String = "integer" ' currently must be integer type (must change field creation code in createattribute

        pOriginaljuncFlagsList.Reset()
        pOriginalBarriersList.Reset()

        Dim stop_check As Boolean = False
        Dim orderLoop As Integer = 0
        Dim m, p, j, k As Integer
        Dim iFCID, iFID, iEID, iSubID, iEID_p, bEID, bFCID, bFID, fEID, bSubID, sOutID, f_sOutID, f_siOutEID, keepEID, endEID As Integer
        Dim sBarrierType, f_sType, sFlagCheck As String
        Dim flagOverBarrier As Boolean = False
        Dim dBarrierPerm As Double
        Dim bBranchJunction As Boolean
        Dim pNextOriginalJuncFlagGEN As IEnumNetEIDBuilderGEN = New EnumNetEIDArray
        Dim pNextOriginalJuncFlag, pNextJunctions, pNodesSet, pBarriersList, pNextTraceBarrierEIDs, pFlowEndJunctionsPer, pFlowEndEdgesPer As IEnumNetEID
        Dim pFilteredBranchJunctionsList, pOriginalEdgeFlagsList, pFilteredSourceJunctionsList, pFlowEndEdges, pResultJunctionsFiltered As IEnumNetEID
        Dim pFirstJunctionBarriers, pFirstEdgeBarriers, pAllFlowEndBarriers, pResultEdges, pResultJunctions, pFlagEnumNetEID As IEnumNetEID
        Dim pNextJunctionsGEN As IEnumNetEIDBuilderGEN = New EnumNetEIDArray
        Dim pNodesSetGEN As IEnumNetEIDBuilderGEN = New EnumNetEIDArray
        Dim pNoSourceFlowEndsTemp As IEnumNetEIDBuilderGEN = New EnumNetEIDArray
        Dim pNextTraceBarrierEIDGEN As IEnumNetEIDBuilderGEN = New EnumNetEIDArray
        Dim pOriginalEdgeFlagsListGEN As IEnumNetEIDBuilderGEN = New EnumNetEIDArray
        Dim pOriginalBarriersListGEN As IEnumNetEIDBuilderGEN = New EnumNetEIDArray
        Dim pFilteredBranchJunctionsGEN As IEnumNetEIDBuilderGEN = New EnumNetEIDArray
        Dim pFilteredSourceJunctionsGEN As IEnumNetEIDBuilderGEN = New EnumNetEIDArray
        Dim pAllFlowEndBarriersGEN As IEnumNetEIDBuilderGEN = New EnumNetEIDArray
        Dim pTotalResultsJunctionsGEN As IEnumNetEIDBuilderGEN = New EnumNetEIDArray
        Dim pTotalResultsEdgesGEN As IEnumNetEIDBuilderGEN = New EnumNetEIDArray
        Dim pFlagEnumNetEIDGEN As IEnumNetEIDBuilderGEN = New EnumNetEIDArray
        Dim pResultJunctionsFilteredGEN As IEnumNetEIDBuilderGEN = New EnumNetEIDArray

        Dim pEnumNetEIDBuilder As IEnumNetEIDBuilder
        Dim pNetworkAnalysisExtResults As INetworkAnalysisExtResults
        Dim pNetworkAnalysisExtBarriers As INetworkAnalysisExtBarriers
        Dim pNetworkAnalysisExtFlags As INetworkAnalysisExtFlags
        Dim lBarrierAndSinkEIDs As List(Of BarrAndBarrEIDAndSinkEIDs) = New List(Of BarrAndBarrEIDAndSinkEIDs)
        Dim pBarrierAndSinkEIDs As New BarrAndBarrEIDAndSinkEIDs(Nothing, Nothing, Nothing, Nothing)
        Dim lBarrierIDs As List(Of BarrierIDObj) = New List(Of BarrierIDObj)
        Dim pNetElements As INetElements
        Dim pFlagDisplay As IFlagDisplay
        Dim pSymbol As ISymbol
        Dim pJuncFlagDisplay As IJunctionFlagDisplay
        Dim pFlagSymbol, pBarrierSymbol As ISimpleMarkerSymbol
        Dim pSimpleMarkerSymbol As ISimpleMarkerSymbol = New SimpleMarkerSymbol
        Dim pRgbColor As IRgbColor = New RgbColor
        Dim pNetworkAnalysisExt As INetworkAnalysisExt
        Dim pNetwork As INetwork
        Dim pTraceFlowSolver As ITraceFlowSolver
        Dim eFlowElements As esriFlowElements
        Dim pIDAndType As New IDandType(Nothing, Nothing)
        Dim bUpHab, bPathDownHab, bSourceJunction As Boolean
        Dim pSelectionSet As ISelectionSet
        Dim lSelectAndWriteFIPEXEIDObject As List(Of SelectAndUpdateFeaturesObject) = New List(Of SelectAndUpdateFeaturesObject)
        Dim lSelectAndWriteFIPEXHGSegmentObject As List(Of SelectAndUpdateFeaturesObject) = New List(Of SelectAndUpdateFeaturesObject)

        Dim pMap As IMap
        Dim pMxDoc As IMxDocument
        Dim pFLyrSlct As IFeatureLayer
        Dim pDoc As IDocument = My.ArcMap.Application.Document
        pMxDoc = CType(pDoc, IMxDocument)
        pMap = pMxDoc.FocusMap
        pMap.ClearSelection()

        pNetworkAnalysisExt = CType(m_UtilityNetworkAnalysisExt, INetworkAnalysisExt)
        pNetworkAnalysisExtFlags = CType(m_UtilityNetworkAnalysisExt, INetworkAnalysisExtFlags)
        pNetworkAnalysisExtBarriers = CType(m_UtilityNetworkAnalysisExt, INetworkAnalysisExtBarriers)
        pNetworkAnalysisExtResults = CType(m_UtilityNetworkAnalysisExt, INetworkAnalysisExtResults)
        Dim pGeometricNetwork As IGeometricNetwork = pNetworkAnalysisExt.CurrentNetwork
        pGeometricNetwork = pNetworkAnalysisExt.CurrentNetwork
        pNetwork = pGeometricNetwork.Network
        pNetElements = CType(pNetwork, INetElements)
        Dim pMxDocument As IMxDocument = CType(My.ArcMap.Application.Document, IMxDocument)
        Dim pActiveView As IActiveView = CType(pMap, IActiveView)

        GetSymbols(pFlagSymbol, pBarrierSymbol)

        ' ******** NO EDGE FLAG SUPPORT YET *********
        Dim i As Integer = 0
        For i = 0 To pNetworkAnalysisExtFlags.EdgeFlagCount - 1
            ' Use the bFlagDisplay to retrieve EIDs of the Edge flags for later
            pFlagDisplay = CType(pNetworkAnalysisExtFlags.EdgeFlag(i), IFlagDisplay)
            bEID = pNetElements.GetEID(pFlagDisplay.FeatureClassID, pFlagDisplay.FID, pFlagDisplay.SubID, esriElementType.esriETEdge)
            pOriginalEdgeFlagsListGEN.Add(bEID)
            'pOriginalEdgeFlagsList(i) = bEID
        Next

        ' QI to and get an array interface that has 'count' and 'next' methods
        pOriginalEdgeFlagsList = CType(pOriginalEdgeFlagsListGEN, IEnumNetEID)

        ' ************ MAIN LOOP *************
        ' to do - expand loop for each user-set flag
        ' For i = 0 To pOriginaljuncFlagsList.Count - 1
        fEID = pOriginaljuncFlagsList.Next()
        ' Query the corresponding user ID's to the element ID

        pNextOriginalJuncFlagGEN = New EnumNetEIDArray
        pNextOriginalJuncFlagGEN.Add(fEID)
        pNextOriginalJuncFlag = CType(pNextOriginalJuncFlagGEN, IEnumNetEID)

        pNodesSet = pNextOriginalJuncFlag

        orderLoop = 0
        stop_check = False
        ' main inner loop for DFS
        Do Until stop_check = True

            orderLoop = orderLoop + 1

            pNextJunctionsGEN = Nothing
            pNextJunctionsGEN = New EnumNetEIDArray

            ' ############## LOOP TO RUN TRACES FOR THIS 'ORDER' ###############
            pNodesSet.Reset()

            ' main loop for each graph 'order'/ 'generation' level
            For j = 0 To pNodesSet.Count - 1

                iEID = pNodesSet.Next()

                ' ============================
                ' ====== FILTER BARRIERS =====
                ' ============================
                k = 0
                pOriginalBarriersList.Reset()
                pNextTraceBarrierEIDGEN = New EnumNetEIDArray

                For k = 0 To pOriginalBarriersList.Count - 1
                    bEID = pOriginalBarriersList.Next()

                    ' Keep if it won't conflict (not on top of flag)
                    If bEID <> iEID Then
                        pNextTraceBarrierEIDGEN.Add(bEID)
                    End If
                Next

                'QI to get 'next' and 'count'
                pNextTraceBarrierEIDs = Nothing
                pNextTraceBarrierEIDs = CType(pNextTraceBarrierEIDGEN, IEnumNetEID)

                ' ===========================
                ' ====== SET BARRIERS  ======
                ' ============================
                k = 0
                pNextTraceBarrierEIDs.Reset()
                pNetworkAnalysisExtBarriers.ClearBarriers()

                For k = 0 To pNextTraceBarrierEIDs.Count - 1
                    bEID = pNextTraceBarrierEIDs.Next()
                    pNetElements.QueryIDs(bEID, esriElementType.esriETJunction, bFCID, bFID, bSubID)

                    ' Display the barriers as a JunctionFlagDisplay type
                    pFlagDisplay = New JunctionFlagDisplay
                    pSymbol = CType(pBarrierSymbol, ISymbol)
                    With pFlagDisplay
                        .FeatureClassID = bFCID
                        .FID = bFID
                        .Geometry = pGeometricNetwork.GeometryForJunctionEID(bEID)
                        .Symbol = pSymbol
                    End With

                    ' Add the flags to the logical network
                    pJuncFlagDisplay = CType(pFlagDisplay, IJunctionFlagDisplay)
                    pNetworkAnalysisExtBarriers.AddJunctionBarrier(pJuncFlagDisplay)
                Next

                ' ===============================
                ' ========== SET FLAG ===========
                ' ===============================
                pNetworkAnalysisExtFlags.ClearFlags()
                pNetElements.QueryIDs(iEID, esriElementType.esriETJunction, iFCID, iFID, iSubID)

                ' Display the flags as a JunctionFlagDisplay type
                pFlagDisplay = New JunctionFlagDisplay
                pSymbol = CType(pFlagSymbol, ISymbol)
                With pFlagDisplay
                    .FeatureClassID = iFCID
                    .FID = iFID
                    .Geometry = pGeometricNetwork.GeometryForJunctionEID(iEID)
                    .Symbol = pSymbol
                End With

                ' Add the flags to the logical network
                pJuncFlagDisplay = CType(pFlagDisplay, IJunctionFlagDisplay)
                pNetworkAnalysisExtFlags.AddJunctionFlag(pJuncFlagDisplay)


                ' ##################################
                ' ########## RUN TRACE UP ##########
                ' ##################################
                ' find the upstream nodes to use in the next loop
                'prepare the network solver, exit if problem
                pTraceFlowSolver = TraceFlowSolverSetup()
                If pTraceFlowSolver Is Nothing Then
                    System.Windows.Forms.MessageBox.Show("Could not set up the network. Check that there is a network loaded.", _
                                                         "TraceFlowSolver setup error.")
                    bContinue = False
                    Exit Sub
                End If

                eFlowElements = esriFlowElements.esriFEJunctionsAndEdges
                pFlowEndJunctionsPer = Nothing
                pFlowEndEdgesPer = Nothing
                pNetworkAnalysisExtResults.ClearResults()

                'Return the features stopping the trace
                ' (sDirection always 'up' atm - GO Sep 2021)
                pTraceFlowSolver.FindFlowEndElements(esriFlowMethod.esriFMUpstream, eFlowElements, pFlowEndJunctionsPer, pFlowEndEdgesPer)


                ' =======================================
                ' =====  FILTER FLOW END ELEMENTS =======
                ' =======================================
                pFlowEndJunctionsPer.Reset()
                pAllFlowEndBarriers = CType(pAllFlowEndBarriersGEN, IEnumNetEID)
                k = 0

                For k = 0 To pFlowEndJunctionsPer.Count - 1
                    keepEID = True
                    endEID = pFlowEndJunctionsPer.Next()
                    m = 0

                    For m = 0 To pAllFlowEndBarriers.Count - 1
                        If endEID = pAllFlowEndBarriers.Next() Then
                            keepEID = False ' set false if already on master list
                        End If
                    Next

                    ' =============================
                    ' =====  REMOVE SOURCES =======
                    If sType = "segment" Then
                        bSourceJunction = IsSourceJunction(m_UtilityNetworkAnalysisExt, endEID)
                        If bSourceJunction = True Then
                            keepEID = False
                        End If
                    End If

                    ' filter out the flag if it's returned as a flow end element
                    ' why is this happening?? - GO Sep 2021
                    If endEID = iEID Then
                        keepEID = False
                    End If

                    If keepEID = True Then
                        ' delete this queryID
                        pNetElements.QueryIDs(endEID, esriElementType.esriETJunction, iFCID, iFID, iSubID)
                        'MsgBox("Keeping EID, FCID, FID for next order loop: " + Str(endEID) + ", " + Str(iFCID) + ", " + Str(iFID))
                        pAllFlowEndBarriersGEN.Add(endEID) ' to prevent infinite loop problem
                        pNextJunctionsGEN.Add(endEID) 'reset each order loop

                    End If
                Next

                ' ===========================================
                ' ===== RETURN UPSTREAM FLOW ELEMENTS =======
                ' ===========================================
                pNetworkAnalysisExtResults.ClearResults()

                ' downstream trace if id'ing subsegments
                If sType = "segment" Then
                    pTraceFlowSolver.FindFlowElements(esriFlowMethod.esriFMUpstream, eFlowElements, pResultJunctions, pResultEdges)
                Else ' if 'subsegment'
                    pTraceFlowSolver.FindFlowElements(esriFlowMethod.esriFMDownstream, eFlowElements, pResultJunctions, pResultEdges)
                End If

                If pResultJunctions Is Nothing Then
                    ' junctions were not returned -- create an empty enumeration
                    pEnumNetEIDBuilder = New EnumNetEIDArray
                    pEnumNetEIDBuilder.Network = pNetworkAnalysisExt.CurrentNetwork.Network
                    pEnumNetEIDBuilder.ElementType = esriElementType.esriETJunction
                    pResultJunctions = CType(pEnumNetEIDBuilder, IEnumNetEID)
                End If

                If pResultEdges Is Nothing Then
                    ' edges were not returned -- create an empty enumeration
                    pEnumNetEIDBuilder = New EnumNetEIDArray
                    pEnumNetEIDBuilder.Network = pNetworkAnalysisExt.CurrentNetwork.Network
                    pEnumNetEIDBuilder.ElementType = esriElementType.esriETEdge
                    pResultEdges = CType(pEnumNetEIDBuilder, IEnumNetEID)
                End If

                ' ======= FILTER THE FLAG FROM RETURNED FEATURES =======
                ' filter out the flag junction from results returned from upstream trace
                ' GO sep 2021 - don't want to attribute the 'segment' of the downstream-most node based on it's own EID (want the dstream EID)
                '             - but will allow geonet junctions that don't serve as FIPEX nodes to get segment labels
                k = 0
                pResultJunctions.Reset()
                pResultJunctionsFilteredGEN = New EnumNetEIDArray
                For k = 0 To pResultJunctions.Count - 1
                    fEID = pResultJunctions.Next()
                    If fEID <> iEID Then
                        pResultJunctionsFilteredGEN.Add(fEID)
                    End If
                Next
                pResultJunctionsFiltered = Nothing
                pResultJunctionsFiltered = CType(pResultJunctionsFilteredGEN, IEnumNetEID)
                pResultJunctionsFiltered.Reset()

                pNetworkAnalysisExtResults.ClearResults()

                ReturnSelectionSetList(iEID, sFieldName, lSelectAndWriteFIPEXHGSegmentObject, pNetworkAnalysisExtResults, _
                                       pNetworkAnalysisExt, pResultJunctionsFiltered, pResultEdges)

                ' ======= SAVE FOR HIGHLIGHTING =======
                ' Get results to display as highlights at end of sub
                ' GO Sep 2021 - not currently using this
                pResultJunctions.Reset()
                k = 0
                For k = 0 To pResultJunctions.Count - 1
                    pTotalResultsJunctionsGEN.Add(pResultJunctions.Next())
                Next
                pResultEdges.Reset()
                For k = 0 To pResultEdges.Count - 1
                    pTotalResultsEdgesGEN.Add(pResultEdges.Next())
                Next

            Next 'node

            pNextJunctions = Nothing
            pNextJunctions = CType(pNextJunctionsGEN, IEnumNetEID)
            pNextJunctions.Reset()

            ' not sure it's necessary to redeclare an empty object for junctions but will do
            pNodesSetGEN = New EnumNetEIDArray
            pNodesSet = Nothing
            pNodesSet = CType(pNodesSetGEN, IEnumNetEID)
            pNodesSet = pNextJunctions
            pNodesSet.Reset()

            pNextJunctionsGEN = Nothing
            pNextJunctionsGEN = New EnumNetEIDArray

            If pNodesSet.Count = 0 Or pNodesSet Is Nothing Then
                'MsgBox("Ending BFS after x loops: " + Str(orderLoop))
                stop_check = True
            End If


        Loop ' 
        ' Next flag

        pMap.ClearSelection()

        'MsgBox("The number of selectionsets stored for flags (number of flags): " + Str(lSelectAndWriteFIPEXEIDObject.Count))
        'MsgBox("The number of selectionsets stored for upstream features (for segment labelling): " + Str(lSelectAndWriteFIPEXHGSegmentObject.Count))

        'Values used can be integer, double, or string (to do: must be int as of Sep 2021, need to change createattribute sub to be flexible)
        UpdateAttributesBatch(lSelectAndWriteFIPEXHGSegmentObject, sValueType)

        pActiveView.Refresh()

    End Sub

    Public Shared Function IsSourceJunction(ByRef m_UtilityNetworkAnalysisExt As ESRI.ArcGIS.EditorExt.IUtilityNetworkAnalysisExt,
                                            ByRef iThisEID As Integer)
        ' Sep 2021
        ' extracted code from FindBranchSourceJunction
        ' returns true if element is a 'source' (no upstream edges)

        Dim iEdges, iEdgeEID, iFromEID, iToEID, k, iFlowDir, iNextEID, iUpstreamEdges As Integer
        Dim bUpstream As Boolean = True
        Dim bIsSource As Boolean = False
        Dim pNetTopology As INetTopology
        Dim pNetwork As INetwork
        Dim pBarriersListWithBranches As IEnumNetEIDBuilderGEN
        Dim pNetElements As INetElements
        Dim pNetworkAnalysisExt As INetworkAnalysisExt
        Dim pUtilityNetwork As IUtilityNetwork ' for flow direction relative to dig. direction
        pNetworkAnalysisExt = CType(m_UtilityNetworkAnalysisExt, INetworkAnalysisExt)
        Dim pGeometricNetwork As IGeometricNetwork = pNetworkAnalysisExt.CurrentNetwork
        pGeometricNetwork = pNetworkAnalysisExt.CurrentNetwork
        pUtilityNetwork = CType(pGeometricNetwork.Network, IUtilityNetwork)
        pNetwork = pGeometricNetwork.Network
        pNetElements = CType(pNetwork, INetElements)
        pNetTopology = CType(pNetwork, INetTopology)


        Dim pForwardStar As IForwardStar
        Dim pForwardStarGEN As IForwardStarGEN

        iEdges = pNetTopology.GetAdjacentEdgeCount(iThisEID)

        ' if adjacent edges to element then for each get the EID and check if upstream
        If iEdges > 0 Then
            For k = 0 To iEdges - 1

                iEdgeEID = 0
                iFromEID = 0
                iToEID = 0

                pNetTopology.GetAdjacentEdge(iThisEID, k, iEdgeEID, True)

                ' Problem is that bOrientation (and other nettopology directions) are digitization direction
                ' no way easily to determine the flow direction of the edges
                ' need to:
                ' get the adjacent edge
                ' determine digitization direction relative to iThisEID
                ' determine flow direction relative to digization direction
                ' is it upstream?

                ' IUtilityNetworkGEN.GetFlowDirection(EID) should get flow direction but it is relative to digitization
                ' https://desktop.arcgis.com/en/arcobjects/latest/net/webframe.htm#esriFlowDirection.htm
                ' iFlowDir = 1 is with digitization direction, 2 = against digitization direction

                pNetTopology.GetFromToJunctionEIDs(iEdgeEID, iFromEID, iToEID)
                iFlowDir = pUtilityNetwork.GetFlowDirection(iEdgeEID)

                If iFromEID = iThisEID Then

                    iNextEID = iToEID
                    If iFlowDir = 1 Then
                        bUpstream = False
                    ElseIf iFlowDir = 2 Then
                        'MsgBox("Debug2020: Edge " & Str(iEdgeEID) & " is upstream of node " & Str(iThisEID))
                        bUpstream = True
                    Else
                        'MsgBox("Debug2020: Edge flow direction is unitialized or indeterminate. For Distance-limited analysis all edges must have determinate flow direction. ")
                        'MsgBox("Debug2020: # The esriflowdirection for edge EID " & Str(iEdgeEID) & " :" & Str(iFlowDir))
                        bUpstream = False
                    End If

                ElseIf iToEID = iThisEID Then
                    iNextEID = iFromEID
                    If iFlowDir = 1 Then
                        bUpstream = True
                    ElseIf iFlowDir = 2 Then
                        bUpstream = False
                    Else
                        bUpstream = False
                    End If
                Else
                    MsgBox("Debug2021: No match?? Issue with flowdir check: " & Str(iThisEID) & " iFromEID: " & Str(iFromEID) & "iToEID: " & Str(iToEID))
                End If

                If bUpstream = True Then
                    iUpstreamEdges = iUpstreamEdges + 1
                End If
            Next
        End If ' adjacent edges > 0

        If iUpstreamEdges > 0 Then
            bIsSource = False
        ElseIf iUpstreamEdges = 0 Then
            bIsSource = True
        End If

        Return bIsSource
    End Function
    Public Shared Sub FindBranchSourceJunctions(ByRef pOriginalBarriersListSaved As IEnumNetEID,
                                                ByRef pNewBarriersList As IEnumNetEID,
                                                ByRef pOriginaljuncFlagsList As IEnumNetEID,
                                                ByRef pNewBarriersListGEN As IEnumNetEIDBuilderGEN,
                                                ByRef m_UtilityNetworkAnalysisExt As ESRI.ArcGIS.EditorExt.IUtilityNetworkAnalysisExt,
                                                ByRef pSourceListGEN As IEnumNetEIDBuilderGEN,
                                                ByRef pBranchListGEN As IEnumNetEIDBuilderGEN)

        ' return an object that can be used to select and assign values to attributes
        ' ######################### FIND ALL BRANCH JUNCTIONS ##############
        ' Created Aug 2020
        ' Ported over to Sharessubs May  2021
        ' Purpose: testing, debugging of finding all junctions representing 
        '          a branch _in addition_ to the junctions representing 
        '          sinks or barriers. Create list of junctions that will 
        '          be treated like 'barriers' for purpose of connectivity 
        '          table generation; amend list of original junction barriers.
        '          the 'original' in object name should be changed since the
        '          actual originally set barriers are stored in pOriginalBarriersListSaved
        ' Logic: 
        ' If 'Advanced Connectivity is checked' 
        ' Clear All Flags and Barriers
        ' For Each Original Flag
        ' Set Flag 
        ' Trace Upstream and find neighbour junction
        ' If more than one upstream immediate junction 
        ' Save current flag junction EID
        ' Move flag to next 'order' and repeat above
        'MsgBox("Debug2020: bAdvConnectTab is set " & Str(bAdvConnectTab))
        Dim pNetTopology As INetTopology
        Dim iEdges, fEID, iFCID, iFID, iSubID As Integer
        Dim iEdgeEID As Integer = 0
        Dim iFromEID As Integer = 0 ' referring to dig. direction
        Dim iToEID As Integer = 0
        Dim iThisEID As Integer = 0
        Dim iNextEID As Integer = 0
        Dim iEID_j As Integer = 0
        Dim iEID_p As Integer = 0
        Dim iEID_k As Integer = 0
        Dim iLastEID As Integer = 0
        Dim iActualFromEID As Integer = 0 ' for actual flow direction
        Dim iActualToEID As Integer = 0
        Dim iFlowDir As Integer = 0 ' esriFlowDirection 
        Dim bOrientation As Boolean = False
        Dim bUpstream As Boolean = False
        Dim bMatch As Boolean = False
        Dim bBranchJunction As Boolean = False
        Dim bSourceJunction As Boolean = False
        Dim iUpstreamEdges As Integer = 0
        Dim pAdjacentEdges, pBranchJunctions, pSourceJunctions, pFilteredBranchJunctionsList, pFilteredSourceJunctionsList As IEnumNetEID
        Dim pAdjacentEdgesGEN, pBranchJunctionsGEN, pSourceJunctionsGEN As IEnumNetEIDBuilderGEN
        Dim pFilteredBranchJunctionsGEN As IEnumNetEIDBuilderGEN = New EnumNetEIDArray
        Dim pFilteredSourceJunctionsGEN As IEnumNetEIDBuilderGEN = New EnumNetEIDArray

        ' next junctions upstream direction
        Dim pNextJunctions, pJunctions As IEnumNetEID
        Dim pJunctionsGEN, pNextJunctionsGEN As IEnumNetEIDBuilderGEN
        Dim iUpstreamJunctionCount As Integer = 0
        Dim pUtilityNetwork As IUtilityNetwork ' need to get flow direction relative to dig. direction
        Dim pForwardStar As IForwardStar
        Dim pForwardStarGEN As IForwardStarGEN
        Dim pNetEdge As INetworkEdge
        Dim iNetEdgeDirection, sUserID As Integer ' the direction of flow relative to digitized direction

        ' May 2021
        Dim pNetwork As INetwork
        Dim pBarriersListWithBranches As IEnumNetEIDBuilderGEN
        Dim pNetElements As INetElements
        Dim pNetworkAnalysisExt As INetworkAnalysisExt
        pNetworkAnalysisExt = CType(m_UtilityNetworkAnalysisExt, INetworkAnalysisExt)
        Dim pGeometricNetwork As IGeometricNetwork = pNetworkAnalysisExt.CurrentNetwork
        pGeometricNetwork = pNetworkAnalysisExt.CurrentNetwork
        pNetwork = pGeometricNetwork.Network
        pNetElements = CType(pNetwork, INetElements)
        Dim pIDAndType As New IDandType(Nothing, Nothing)

        pOriginaljuncFlagsList.Reset()
        pBranchJunctionsGEN = New EnumNetEIDArray
        pSourceJunctionsGEN = New EnumNetEIDArray

        ' for each flag the user has set 
        For i = 0 To pOriginaljuncFlagsList.Count - 1

            pNetTopology = CType(pNetwork, INetTopology)
            fEID = pOriginaljuncFlagsList.Next()

            ' get user label for result form
            pNetElements.QueryIDs(fEID, esriElementType.esriETJunction, iFCID, iFID, iSubID)

            ' Use Do loop until no upstream neighbours are found
            pJunctionsGEN = New EnumNetEIDArray
            pJunctionsGEN.Add(fEID) ' initialize this object before first loop
            pJunctions = CType(pJunctionsGEN, IEnumNetEID)
            iUpstreamJunctionCount = pJunctions.Count ' initialize for first loop
            pUtilityNetwork = CType(pGeometricNetwork.Network, IUtilityNetwork)
            pForwardStarGEN = pUtilityNetwork.CreateForwardStar(True, Nothing, Nothing, Nothing, Nothing)
            pForwardStar = CType(pForwardStarGEN, IForwardStar)

            Do Until iUpstreamJunctionCount = 0

                pJunctions.Reset()
                pNextJunctionsGEN = New EnumNetEIDArray

                For j = 0 To pJunctions.Count - 1
                    iThisEID = pJunctions.Next()
                    iUpstreamEdges = 0

                    'pForwardStar.FindAdjacent(0, iThisEID, iEdges)
                    iEdges = pNetTopology.GetAdjacentEdgeCount(iThisEID)
                    'MsgBox("Debug2020: # edge neighbours found for node EID " & Str(iThisEID) & "using pForwardStar: " & Str(iEdges))

                    If iEdges > 0 Then
                        'If pNetTopology.GetAdjacentEdgeCount(iThisEID) > 0 Then
                        'pEnumNetEIDBuilder.Add(lEID)
                        'MsgBox("Debug2020: neighbours found for flag " & Str(fEID))
                        ' if edge count is greater than 2 (one uo one down)
                        ' add current junction to list of branch junctions
                        ' for each edge find adjacent junctions
                        'iEdges = pNetTopology.GetAdjacentEdgeCount(iThisEID)
                        'MsgBox("Debug2020: # edge neighbours found for node EID " & Str(iThisEID) & "using pNetTopology: " & Str(iEdges))

                        For k = 0 To iEdges - 1

                            iEdgeEID = 0
                            iFromEID = 0
                            iToEID = 0

                            'MsgBox("Debug2020: # This junction  EID: " & Str(iThisEID))
                            'MsgBox("Debug2020: edgeEID initialized to: " & Str(iEdgeEID))
                            pNetTopology.GetAdjacentEdge(iThisEID, k, iEdgeEID, True)
                            'MsgBox("Debug2020: neighbour edge EID found using pNetTopology: " & Str(iEdgeEID))

                            'pForwardStar.QueryAdjacentEdge(k, iEdgeEID, bOrientation, Nothing)
                            ' MsgBox("Debug2020: neighbour edge EID found using pForwardStar: " & Str(iEdgeEID))
                            'MsgBox("Debug2020: is neighbour edge " & Str(iEdgeEID) & " found using pForwardStar upstream? " & Str(bOrientation))

                            ' Problem is that bOrientation (and other nettopology directions) are digitization direction
                            ' no way easily to determine the flow direction of the edges
                            ' need to:
                            ' get the adjacent edge
                            ' determine digitization direction relative to iThisEID
                            ' determine flow direction relative to digization direction
                            ' is it upstream?

                            ' IUtilityNetworkGEN.GetFlowDirection(EID) should get flow direction but it is relative to digitization
                            ' https://desktop.arcgis.com/en/arcobjects/latest/net/webframe.htm#esriFlowDirection.htm
                            ' 1 = with digitization direction, 2 = against digitization direction
                            'pUtilityNetwork.
                            'iToEID = iThisEID
                            'MsgBox("Debug2020: # The initialized next junction EID: " & Str(iNextJunctionEID))
                            pNetTopology.GetFromToJunctionEIDs(iEdgeEID, iFromEID, iToEID)
                            'MsgBox("Debug2020: # The returned 'to' node using PNetTopology for edge " & Str(iEdgeEID) & " :" & Str(iToEID))
                            'MsgBox("Debug2020: # The returned 'from' node using PNetTopology for edge " & Str(iEdgeEID) & " :" & Str(iFromEID))
                            'pForwardStar.QueryAdjacentJunction(k, iFromEID, Nothing)
                            'MsgBox("Debug2020: # The neighbour node using pForwardStar for edge " & Str(iEdgeEID) & " :" & Str(iFromEID))

                            iFlowDir = pUtilityNetwork.GetFlowDirection(iEdgeEID)
                            'MsgBox("Debug2020: # The esriflowdirection for edge " & Str(iEdgeEID) & " :" & Str(iFlowDir))

                            If iFromEID = iThisEID Then
                                iNextEID = iToEID
                                If iFlowDir = 1 Then
                                    'MsgBox("Debug2020: Edge " & Str(iEdgeEID) & " is downstream of node " & Str(iThisEID))
                                    bUpstream = False
                                ElseIf iFlowDir = 2 Then
                                    'MsgBox("Debug2020: Edge " & Str(iEdgeEID) & " is upstream of node " & Str(iThisEID))
                                    bUpstream = True
                                Else
                                    MsgBox("Debug2020: Edge flow direction is unitialized or indeterminate. For Distance-limited analysis all edges must have determinate flow direction. ")
                                    MsgBox("Debug2020: # The esriflowdirection for edge EID " & Str(iEdgeEID) & " :" & Str(iFlowDir))
                                    'pNetElements.QueryIDs(iEID, esriElementType.esriETJunction, _
                                    '  iFCID, _
                                    '  iFID, _
                                    '  iSubID)
                                    bUpstream = False
                                End If

                            ElseIf iToEID = iThisEID Then
                                iNextEID = iFromEID
                                If iFlowDir = 1 Then
                                    bUpstream = True
                                ElseIf iFlowDir = 2 Then
                                    bUpstream = False
                                Else
                                    bUpstream = False
                                End If
                            Else
                                MsgBox("Debug2020: No match?? iThisEID: " & Str(iThisEID) & " iFromEID: " & Str(iFromEID) & "iToEID: " & Str(iToEID))
                            End If

                            ' add the upstream EID to a list of EIDs unless we are 
                            ' Note here would be the place to identify braids, if that is desired
                            '    Do this after this loop by checking if any duplicates in list of 
                            '    upstream EIDs
                            If bUpstream = True Then
                                pNextJunctionsGEN.Add(iNextEID)
                                iUpstreamEdges = iUpstreamEdges + 1
                            End If
                        Next

                    End If ' upstream edges > 0
                    ' Add to list of branching junctions

                    If iUpstreamEdges > 1 Then
                        pBranchJunctionsGEN.Add(iThisEID)
                        ' going to add this EID to list and remove duplicates later
                        'pOriginalBarriersListGEN.Add(iThisEID)
                    ElseIf iUpstreamEdges = 0 Then
                        pSourceJunctionsGEN.Add(iThisEID)
                    End If

                    'iLastEID = iThisEID ' track last loops EID 

                Next 'junction in 'order
                pNextJunctions = Nothing
                pNextJunctions = CType(pNextJunctionsGEN, IEnumNetEID)
                pNextJunctions.Reset()

                ' not sure it's necessary to redeclare an empty object for junctions but will do
                pJunctionsGEN = New EnumNetEIDArray
                pJunctions = Nothing
                pJunctions = CType(pJunctionsGEN, IEnumNetEID)
                pJunctions = pNextJunctions
                pNextJunctionsGEN = New EnumNetEIDArray
                iUpstreamJunctionCount = pJunctions.Count()

            Loop

            ' merge the original barriers list and the branch junction list
            ' careful because the original barriers list likely has other objects
            ' linked to it that contain permeability, ID, etc etc
            ' remove duplicates

        Next ' original junction flag or sink


        ' For each pBranchJunctions add it to the original barrier list generator
        pBranchJunctions = CType(pBranchJunctionsGEN, IEnumNetEID)

        pBranchJunctions.Reset()
        For j = 0 To pBranchJunctions.Count - 1
            iEID_j = pBranchJunctions.Next()

            ' make sure it's not a duplicate 
            '(if user has placed flag on branch junction, then omit the branch junction)
            pNewBarriersList.Reset()
            bMatch = False

            For k = 0 To pNewBarriersList.Count - 1
                iEID_k = pNewBarriersList.Next()
                If iEID_j = iEID_k Then
                    bMatch = True
                End If
            Next
            If bMatch = False Then
                ' add branch junction to user-set barrier list 
                ' this list arrives pre-populated with the user set barriers - GO May 2021
                pNewBarriersListGEN.Add(iEID_j)
                ' store filtered (no user-set barrier) branch junction for later
                pFilteredBranchJunctionsGEN.Add(iEID_j)

            End If
        Next

        pFilteredBranchJunctionsList = Nothing
        pFilteredBranchJunctionsList = CType(pFilteredBranchJunctionsGEN, IEnumNetEID)
        pFilteredBranchJunctionsList.Reset()
        pSourceJunctions = CType(pSourceJunctionsGEN, IEnumNetEID)
        pSourceJunctions.Reset()

        For j = 0 To pSourceJunctions.Count - 1
            iEID_j = pSourceJunctions.Next()

            ' make sure it's not a duplicate 
            '(if user has placed flag on branch junction, then omit the branch junction)
            pNewBarriersList.Reset()
            bMatch = False

            For k = 0 To pNewBarriersList.Count - 1
                iEID_k = pNewBarriersList.Next()
                If iEID_j = iEID_k Then
                    bMatch = True

                End If
            Next
            If bMatch = False Then
                ' add branch junction to user-set barrier list 
                pNewBarriersListGEN.Add(iEID_j)
                ' store filtered (no user-set barrier) source junction for later
                pFilteredSourceJunctionsGEN.Add(iEID_j)

            End If
        Next

        pFilteredSourceJunctionsList = Nothing
        pFilteredSourceJunctionsList = CType(pFilteredSourceJunctionsGEN, IEnumNetEID)
        pFilteredSourceJunctionsList.Reset()
        pNewBarriersList = Nothing
        pNewBarriersList = CType(pNewBarriersListGEN, IEnumNetEID)

        pSourceListGEN = pFilteredSourceJunctionsGEN
        pBranchListGEN = pFilteredBranchJunctionsGEN

    End Sub
    Public Shared Sub LabelBranchSourceJunctions(ByRef sAttributeName As String,
                                                 ByRef pOriginalBarriersListSaved As IEnumNetEID,
                                                 ByRef pOriginaljuncFlagsList As IEnumNetEID,
                                                 ByRef pSourceElements As IEnumNetEID,
                                                 ByRef pBranchElements As IEnumNetEID,
                                                 ByRef pUtilityNetworkAnalysisExt As ESRI.ArcGIS.EditorExt.IUtilityNetworkAnalysisExt)
        ' - GO Sep 2021, work in progress
        ' for each flag the user has set 

        ' for each barrier, branch, sink, source node, get the FCID, FID, workspace etc and store it in object
        ' start edit session and attempt to update the attribute

        Dim i, fEID, iFCID, iFID, iSubID As Integer
        Dim pNetworkAnalysisExt As INetworkAnalysisExt
        pNetworkAnalysisExt = CType(pUtilityNetworkAnalysisExt, INetworkAnalysisExt)
        Dim pGeometricNetwork As IGeometricNetwork = pNetworkAnalysisExt.CurrentNetwork
        pGeometricNetwork = pNetworkAnalysisExt.CurrentNetwork
        Dim pNetwork As INetwork
        pNetwork = pGeometricNetwork.Network
        Dim pNetElements As INetElements
        pNetElements = CType(pNetwork, INetElements)
        Dim pNetTopology As INetTopology

        For i = 0 To pOriginalBarriersListSaved.Count - 1

            pNetTopology = CType(pNetwork, INetTopology)
            fEID = pOriginalBarriersListSaved.Next()

            ' get user label for result form
            ' the UserClassID and UserID correspond to the FeatureClassID and OID of the feature.
            '  The UserSubID is the ID of the subelement of the feature.
            pNetElements.QueryIDs(fEID, esriElementType.esriETJunction, iFCID, iFID, iSubID)

            ' store as selection set(s) and then use one edit session to update
            ' see FeatureLayerAndSelectionSet
            'Dim lFeatureLayerAndSelSetLINE As List(Of FeatureLayerAndSelectionSet) = New List(Of FeatureLayerAndSelectionSet)
            'Dim pFeatureLayerAndSelectionSet As FeatureLayerAndSelectionSet = New FeatureLayerAndSelectionSet(Nothing, Nothing)

            ' get the selection set
            'pFeatureSelectionLine = Nothing
            'pFeatureSelectionLine = CType(pFeatureLayerMap, IFeatureSelection)
            'pSelectionSetLine = Nothing ' like to reset these cuz sometimes whothefuckknowswhy it causes issues otherwise
            'pSelectionSetLine = pFeatureSelectionLine.SelectionSet

            'pFeatureLayerAndSelectionSet = New FeatureLayerAndSelectionSet(pFeatureLayerMap, pSelectionSetLine)
            'lFeatureLayerAndSelSetLINE.Add(pFeatureLayerAndSelectionSet)

        Next
    End Sub

    Public Shared Sub CheckCreateAttribute(ByRef sAttributeName As String)
        ' - GO Sep 2021, work in progress
        ' - checks and inserts new attribute into each geonetwork junctions layer

        ' For each junctions layer in the geometric network
        '  check for the attribute
        '  if not there, notify user, attempt to create attribute

    End Sub

    Public Shared Sub ResultsForm2020(ByRef pResultsForm3 As FiPEX_ArcMap_10p4_up_AddIn_dotNet45_2020.frmResults_3,
                                      ByRef lSinkIDandTypes As List(Of SinkandTypes),
                                      ByRef lHabStatsList As List(Of StatisticsObject_2),
                                      ByRef lMetricsObject As List(Of MetricsObject),
                                      ByRef BeginTime As DateTime, ByRef numbarrsnodes As String,
                                      ByRef iOrderNum As Integer, sDirection As String)
        ' col 0 - sink ID
        ' col 1 - sink EID
        ' col 2 - sink node type (barrier / junction )
        ' col 3 - barrier ID
        ' col 4 - barrier EID
        ' col 5 - stat (e.g., perm, DCI)
        ' col 6 - trace type (e.g. upstream)
        ' col 7 - trace subtype (e.g., immediate)
        ' col 8 - class
        ' col 9 - value
        ' col 10 - units
        ' col 11 - hab_dimension (can be length, area)
        ' col 10 - units

        ' Output Form (will replace dockable window)
        'Dim pResultsForm3 As New FiPEX_ArcMap_10p4_up_AddIn_dotNet45_2020.frmResults_3
        Dim iSinkRowIndex As Integer


        ' Set up the table - create columns 
        pResultsForm3.DataGridView1.Columns.Add("SinkID", "Sink User ID")             '0
        pResultsForm3.DataGridView1.Columns.Add("SinkEID", "Sink Net EID")           '1
        pResultsForm3.DataGridView1.Columns.Add("SinkNodeType", "Sink Node Type") '2
        pResultsForm3.DataGridView1.Columns.Add("BarrierID", "Barrier User ID")       '3
        pResultsForm3.DataGridView1.Columns.Add("BarrierEID", "Barrier Net EID")     '4
        pResultsForm3.DataGridView1.Columns.Add("Stat", "Statistic")                 '5
        pResultsForm3.DataGridView1.Columns.Add("TraceType", "Trace Type")       '6
        pResultsForm3.DataGridView1.Columns.Add("TraceSubtype", "Trace Subtype") '7
        pResultsForm3.DataGridView1.Columns.Add("class", "class")               '8
        pResultsForm3.DataGridView1.Columns.Add("value", "value")               '9
        pResultsForm3.DataGridView1.Columns.Add("units", "units")               '10
        pResultsForm3.DataGridView1.Columns.Add("dimension", "dimension")       '11
        pResultsForm3.DataGridView1.Columns.Add("layer", "layer")               '12

        pResultsForm3.DataGridView1.Columns(0).Width = 46
        pResultsForm3.DataGridView1.Columns(1).Width = 46
        pResultsForm3.DataGridView1.Columns(2).Width = 46
        pResultsForm3.DataGridView1.Columns(3).Width = 46
        pResultsForm3.DataGridView1.Columns(4).Width = 46
        pResultsForm3.DataGridView1.Columns(5).Width = 72
        pResultsForm3.DataGridView1.Columns(6).Width = 65
        pResultsForm3.DataGridView1.Columns(7).Width = 65
        pResultsForm3.DataGridView1.Columns(8).Width = 90
        pResultsForm3.DataGridView1.Columns(9).Width = 46
        pResultsForm3.DataGridView1.Columns(10).Width = 46
        pResultsForm3.DataGridView1.Columns(11).Width = 65
        pResultsForm3.DataGridView1.Columns(12).Width = 75


        For i = 0 To lSinkIDandTypes.Count - 1

            'pDataGridViewCellStyle = pResultsForm3.DataGridView1.Rows(iSinkRowIndex).Cells(0).Style
            'pDataGridViewCellStyle = pResultsForm3.DataGridView1.Rows(iSinkRowIndex).Cells(1).Style
            'pDataGridViewCellStyle.Font = New Font(pResultsForm3.DataGridView1.Font.FontFamily, 14, FontStyle.Bold)
            pResultsForm3.DataGridView1.AllowUserToResizeColumns = True
            pResultsForm3.DataGridView1.AllowUserToResizeRows = True
            pResultsForm3.DataGridView1.ColumnHeadersDefaultCellStyle.Font = New Font(pResultsForm3.DataGridView1.Font.FontFamily, 8, FontStyle.Bold)

            For j = 0 To lMetricsObject.Count - 1
                If lMetricsObject(j).SinkEID = lSinkIDandTypes(i).SinkEID Then
                    iSinkRowIndex = pResultsForm3.DataGridView1.Rows.Add()
                    pResultsForm3.DataGridView1.Rows(iSinkRowIndex).Cells(0).Value = lSinkIDandTypes(i).SinkID
                    pResultsForm3.DataGridView1.Rows(iSinkRowIndex).Cells(1).Value = lSinkIDandTypes(i).SinkEID
                    pResultsForm3.DataGridView1.Rows(iSinkRowIndex).Cells(2).Value = lSinkIDandTypes(i).Type
                    pResultsForm3.DataGridView1.Rows(iSinkRowIndex).Cells(3).Value = lMetricsObject(j).ID
                    pResultsForm3.DataGridView1.Rows(iSinkRowIndex).Cells(4).Value = lMetricsObject(j).BarrEID
                    pResultsForm3.DataGridView1.Rows(iSinkRowIndex).Cells(5).Value = lMetricsObject(j).MetricName
                    pResultsForm3.DataGridView1.Rows(iSinkRowIndex).Cells(6).Value = "-"
                    pResultsForm3.DataGridView1.Rows(iSinkRowIndex).Cells(7).Value = "-"
                    pResultsForm3.DataGridView1.Rows(iSinkRowIndex).Cells(8).Value = "-"
                    pResultsForm3.DataGridView1.Rows(iSinkRowIndex).Cells(9).Value = Math.Round(lMetricsObject(j).Metric, 2)
                    pResultsForm3.DataGridView1.Rows(iSinkRowIndex).Cells(10).Value = "-"
                    pResultsForm3.DataGridView1.Rows(iSinkRowIndex).Cells(11).Value = "-"
                    pResultsForm3.DataGridView1.Rows(iSinkRowIndex).Cells(12).Value = "-"
                End If
            Next

            ' for each habitat object item add the stats (hab by class, etc)
            For j = 0 To lHabStatsList.Count - 1
                If lHabStatsList(j).SinkEID = lSinkIDandTypes(i).SinkEID Then
                    iSinkRowIndex = pResultsForm3.DataGridView1.Rows.Add()
                    pResultsForm3.DataGridView1.Rows(iSinkRowIndex).Cells(0).Value = lSinkIDandTypes(i).SinkID
                    pResultsForm3.DataGridView1.Rows(iSinkRowIndex).Cells(1).Value = lSinkIDandTypes(i).SinkEID
                    pResultsForm3.DataGridView1.Rows(iSinkRowIndex).Cells(2).Value = lSinkIDandTypes(i).Type
                    pResultsForm3.DataGridView1.Rows(iSinkRowIndex).Cells(3).Value = lHabStatsList(j).bID
                    pResultsForm3.DataGridView1.Rows(iSinkRowIndex).Cells(4).Value = lHabStatsList(j).bEID
                    pResultsForm3.DataGridView1.Rows(iSinkRowIndex).Cells(5).Value = lHabStatsList(j).LengthOrHabitat
                    pResultsForm3.DataGridView1.Rows(iSinkRowIndex).Cells(6).Value = lHabStatsList(j).Direction
                    pResultsForm3.DataGridView1.Rows(iSinkRowIndex).Cells(7).Value = lHabStatsList(j).TotalImmedPath
                    pResultsForm3.DataGridView1.Rows(iSinkRowIndex).Cells(8).Value = lHabStatsList(j).UniqueClass
                    pResultsForm3.DataGridView1.Rows(iSinkRowIndex).Cells(9).Value = Math.Round(lHabStatsList(j).Quantity, 2)
                    pResultsForm3.DataGridView1.Rows(iSinkRowIndex).Cells(10).Value = lHabStatsList(j).Unit
                    pResultsForm3.DataGridView1.Rows(iSinkRowIndex).Cells(11).Value = lHabStatsList(j).HabitatDimension
                    pResultsForm3.DataGridView1.Rows(iSinkRowIndex).Cells(12).Value = lHabStatsList(j).Layer

                End If
            Next
        Next

        Dim EndTime As DateTime
        EndTime = DateTime.Now

        Dim TotalTime As TimeSpan
        TotalTime = EndTime - BeginTime
        pResultsForm3.lblBeginTime.Text = "Begin Time: " & BeginTime
        pResultsForm3.lblEndtime.Text = "End Time: " & EndTime
        pResultsForm3.lblTotalTime.Text = "Total Time: " & TotalTime.Hours & "hrs " & TotalTime.Minutes & "minutes " & TotalTime.Seconds & "seconds"
        pResultsForm3.lblDirection.Text = "Analysis Direction: " + sDirection
        If iOrderNum <> 99999 Then
            pResultsForm3.lblOrder.Text = "Order of Analysis: " & CStr(iOrderNum)
        Else
            pResultsForm3.lblOrder.Text = "Order of Analysis: Max (all nodes in analysis direction)"
        End If

        If Not numbarrsnodes Is Nothing Then
            pResultsForm3.lblNumBarriers.Text = "Number of Barriers / Nodes Analysed: " & numbarrsnodes
        Else
            pResultsForm3.lblNumBarriers.Text = "Number of Barriers / Nodes Analysed: 1"
        End If
        pResultsForm3.BringToFront()

    End Sub
    Public Shared Sub exclusions2020(ByRef bExclude As Boolean, ByRef pFeature As IFeature, ByRef pFeatureLayer As IFeatureLayer)

        ' =============================================
        ' ============== EXCLUSIONS 2020 ==============
        Dim e_FiPEx__1 As FishPassageExtension
        If e_FiPEx__1 Is Nothing Then
            e_FiPEx__1 = FiPEX_ArcMap_10p4_up_AddIn_dotNet45_2020.FishPassageExtension.GetExtension()
        End If

        Dim iExclusions, j As Integer
        Dim plExclusions As List(Of LayerToExclude) = New List(Of LayerToExclude)
        If e_FiPEx__1.m_bLoaded = True Then ' If there were any extension settings set

            iExclusions = Convert.ToInt32(e_FiPEx__1.pPropset.GetProperty("numExclusions"))
            Dim ExclusionsObj As New LayerToExclude(Nothing, Nothing, Nothing)

            ' match any of the line layers saved in stream to those in listboxes
            If iExclusions > 0 Then
                For j = 0 To iExclusions - 1
                    'sLineLayer = m_FiPEX__1.pPropset.GetProperty("IncLine" + j.ToString) ' get line layer
                    ExclusionsObj = New LayerToExclude(Nothing, Nothing, Nothing)
                    With ExclusionsObj
                        '.Layer = sLineLayer
                        .Layer = Convert.ToString(e_FiPEx__1.pPropset.GetProperty("ExcldLayer" + j.ToString))
                        .Feature = Convert.ToString(e_FiPEx__1.pPropset.GetProperty("ExcldFeature" + j.ToString))
                        .Value = Convert.ToString(e_FiPEx__1.pPropset.GetProperty("ExcldValue" + j.ToString))
                    End With

                    ' add to the module level list
                    plExclusions.Add(ExclusionsObj)
                Next
            End If
        Else
            MsgBox("The FIPEX Options haven't been loaded.  Exiting Exclusions Subroutine. FIPEX code 2642. ")
            Exit Sub
        End If

        Dim x As Integer = 0
        Dim iFieldVal As Integer
        Dim vVal As Object
        Dim sTempVal As String
        For x = 0 To iExclusions - 1
            If pFeatureLayer.Name = plExclusions(x).Layer Then
                ' try to find the field
                iFieldVal = pFeature.Fields.FindField(plExclusions(x).Feature)
                If iFieldVal <> -1 Then
                    Try
                        vVal = pFeature.Value(iFieldVal)
                    Catch ex As Exception
                        MsgBox("Could not convert value found in Exclusions Sub to type 'variant/object'. " + _
                               ". Found in Layer " + pFeatureLayer.Name.ToString + ex.Message)
                        vVal = "nothing"
                    End Try
                    Try
                        sTempVal = vVal.ToString
                    Catch ex As Exception
                        MsgBox("Could not convert value found in Exclusions Sub to type 'string'. " + _
                               ". Found in Layer " + pFeatureLayer.Name.ToString + ex.Message)
                        sTempVal = "nothing"
                    End Try

                    If sTempVal IsNot Nothing Then
                        If sTempVal = plExclusions(x).Value Then
                            bExclude = True
                        End If
                    End If
                End If
            End If
        Next

    End Sub
    Public Shared Sub calculateStatistics_2020(ByRef lHabStatsList As List(Of StatisticsObject_2), _
                                        ByRef lLineLayersFields As List(Of LineLayerToAdd), _
                                        ByRef lPolyLayersFields As List(Of PolyLayerToAdd), _
                                        ByRef ID As String, _
                                        ByRef iEID As Integer, _
                                        ByRef sType As String, _
                                        ByRef f_sOutID As String, _
                                        ByRef f_siOutEID As Integer,
                                        ByVal sHabTypeKeyword As String, _
                                        ByVal sDirection2 As String)

        ' **************************************************************************************
        ' Subroutine:  Calculate Statistics (2) 
        ' Author:       G Oldford
        ' Purpose:     1) intersect other included layers with returned selection
        '                 from the trace.
        '              2) calculate habitat area and length using habitat classes 
        '                 and excluding unwanted features
        '              3) get array (matrix?) of statistics for each habitat class and
        '                 each layer included for habitat classification stats
        '              4) update statistics object and send back to onclick
        ' Keywords:    sHabTypeKeyword - "Total", "Immediate", or "Path"
        '
        '
        ' Notes:
        ' 
        '       Aug, 2020    --> Not sure why passing vars by ref other than lHabStatsList.
        '                        Deleted 'sKeyword' arg (checks if flag on barr or nonbarr) - it was unused. 
        '                        The object returned must differentiate between 'length' and 'habitat' for DD,  
        '                        so changed object by adding a 'LengthOrHabitat' param. 
        '                        Since object returned is used in output tables (not only for DCI) then I must 
        '                        keep all habitat returned (not just either line or poly habitat depending on user
        '                        choice). introduced object to tag the habitat as 'length' or 
        '                        'area' based on whether it is drawn from polygon or line layer. It's probably 
        '                        a better option long-term to use the units as the basis for differentiating 
        '                        whether a quantity returned from a feature represents habitat area or habitat length. 
        '                        In other words, right now the TOC layer type determines whether the habitat extracted from
        '                        the TOC layer is 'area' or 'line'
        ' 
        ' 
        '       Oct 5, 2010  --> Changing this subroutine to a function so it can update the statistics 
        '                  object for habitat statistics (with classes) ONLY. i.e., there will be no 
        '                  other metrics included in this habitat statistics object.
        '                  Added another keyword to say whether this is TOTAL habitat or otherwise (sHabTypeKeyword). 
        '    
        '       Mar 3, 2008  --> only polygon feature layers are intersected.  The function
        '                  checks the config file for included polygons and will intersect any
        '                  network features returned by the trace with the polygons on the list.
        '                  There is probably no reason to have this explicitly for polygons, and
        '                  dividing the 'includes' list into line and polygon categories means that
        '                  the habitat classification also must be divided as such.  This would double
        '                  the number of variables for this process (polygon habitat class layer
        '                  variable, line hab class lyr var, polygon hab class case field var, etc.)
        '                  So since network feature layers are already being returned by the trace,
        '                  they don't need to be intersected.  If we have one 'includes' list that
        '                  contains both polygon and line layers then we need to find out which layers
        '                  in this list are not part of the geometric network, and only intersect these
        '                  features.
        '                  For each includes feature, For each current geometric feature, find match?  Next
        '                  If no match then continue intersection.

        Dim pMxDoc As IMxDocument
        Dim pEnumLayer As IEnumLayer
        Dim pFeatureLayer As IFeatureLayer
        Dim pFeatureSelection As IFeatureSelection
        Dim pFeatureCursor As IFeatureCursor
        Dim pFeature As IFeature
        ' Feb 29 --> There will be a variable number of "included" layers
        '            to use for the habitat classification summary tables.
        '            Each table corresponds to "pages" in the matrix.
        '            Matrix(pages, columns, rows)
        '            Only the farthest right element in a matrix can be
        '            redim "preserved" in VB6 meaning there must be a static
        '            number of columns and pages.  Pages isn't a problem.
        '            They will be the number of layers in the "includes" list
        '            Columns, however, will vary.  This is a problem.  They
        '            will vary between pages of the matrix too which means there
        '            will be empty columns on at least one page if the column count
        '            is different between pages.
        '            Answer to this problem is to avoid the matrix altogether and
        '            update the necessary tables within this function
        Dim e_FiPEx__1 As FishPassageExtension
        If e_FiPEx__1 Is Nothing Then
            e_FiPEx__1 = FiPEX_ArcMap_10p4_up_AddIn_dotNet45_2020.FishPassageExtension.GetExtension()
        End If

        Dim pUID As New UID
        ' Get the pUID of the SelectByLayer command
        'pUID.Value = "{82B9951B-DD63-11D1-AA7F-00C04FA37860}"
        'Dim pGp As IGeoProcessor
        'pGp = New ESRI.ArcGIS.Geoprocessing.GeoProcessor
        Dim pMxDocument As IMxDocument
        Dim pMap As IMap
        Dim i, j, k, m As Integer
        Dim iFieldVal As Integer  ' The field index
        Dim pFields As IFields
        Dim vVar As Object
        Dim pSelectionSet As ISelectionSet
        Dim sTemp As String
        Dim sUnit As String

        ' K REPRESENTS NUMBER OF POSSIBLE HABITAT CLASSES
        '  rows, columns.  ROWS SHOULD BE SET BY NUMBER OF SUMMARY FIELDS
        ' cannot be redimension preserved later
        Dim lHabStatsMatrix As New List(Of HabStatisticsObject)
        Dim pHabStatisticsObject As New HabStatisticsObject(Nothing, Nothing)

        'Dim pFeatureWkSp As IFeatureWorkspace
        Dim pDataStats As IDataStatistics
        Dim pCursor As ICursor
        Dim vFeatHbClsVl As Object ' Feature Habitat Class Value (an object because classes can be numbers or string)
        Dim vTemp As Object
        Dim sFeatClassVal As String
        Dim sMatrixVal As String
        Dim dHabArea, dHabLength As Double
        Dim bClassFound As Boolean
        'For k = 1 To UBound(mHabClassVals, 2) vb6
        Dim classComparer As FindStatsClassPredicate2020
        Dim iStatsMatrixIndex As Integer ' for refining statistics list 
        Dim sClass As String
        Dim vHabTemp As Object
        Dim pDoc As IDocument = My.ArcMap.Application.Document
        pMxDoc = CType(pDoc, IMxDocument)
        pMxDocument = CType(pDoc, IMxDocument)
        pMap = pMxDocument.FocusMap

        ' 2020 - change this two separate objects, lines polygons
        ' layer to hold parameters to send to property
        Dim PolyHabLayerObj As New PolyLayerToAdd(Nothing, Nothing, Nothing, Nothing)
        Dim LineHabLayerObj As New LineLayerToAdd(Nothing, Nothing, Nothing, Nothing, Nothing, Nothing)

        ' object to hold stats to add to list. 
        Dim pHabStatsObject_2 As New StatisticsObject_2(Nothing, Nothing, Nothing, Nothing, Nothing, Nothing, Nothing, Nothing, Nothing, Nothing, Nothing, Nothing, Nothing, Nothing, Nothing)
        Dim sDirection As String
        sDirection = Convert.ToString(e_FiPEx__1.pPropset.GetProperty("direction"))

        ' ================ 2.0 Calculate Area and Length ======================
        ' This next section calculates the area or length of selected features
        ' in the TOC.
        '
        ' PROCESS LOGIC:
        '  1.0 For each Feature Layer in the map
        '  1.1 Filter out any excluded features
        '  1.2 Get a list of all fields in the layer
        '  1.3 Combine the polygon and line layers into one list
        '      2020 - changed this so they are no longer combined
        '  1.4 Prepare the dockable window 
        '    2.0 For each habitat layer in the new list (polygons and lines)
        '      3.0 If there's a match b/w the current layer and habitat layer in list
        '        4.0 then prepare Dockable Window and DBF tables if need be
        '        4.1 Search for the habitat class field in layer
        '        4.2a If the field is found
        '          5.0a If there is a selection set 
        '            6.0a Get the unique values in that field from the selection set
        '            6.1a Loop through unique values and add each to the left column
        '                of a two-column array/matrix to hold statistics
        '            6.2a For each selected feature in the layer
        '              7.0a Get the value in the habitat class field
        '              7.1a For each unique habitat class value in the statistics matrix
        '                8.0a If it matches the value of the class field found in the current feature
        '                  9.0a then add the value of the quantity field in that feature to the
        '                      quantity field for that row in the matrix
        '        4.2b Else if the habitat class field is not found
        '          5.0b If there is a selection set
        '            6.0b For each feature total up stats
        '          5.1b Send output to dockable window

        pUID = New UID
        pUID.Value = "{E156D7E5-22AF-11D3-9F99-00C04F6BC78E}"
        pEnumLayer = pMap.Layers(pUID, True)
        pEnumLayer.Reset()

        ' Look at the next layer in the list
        pFeatureLayer = CType(pEnumLayer.Next(), IFeatureLayer)
        Dim iClassCheckTemp, iLengthField As Integer
        Dim iLoopCount As Integer = 0
        Dim dTempQuan As Double = 0
        Dim dTotalLength As Double = 0
        Dim bExclude As Boolean = False

        Do While Not pFeatureLayer Is Nothing ' these two lines must be separate
            If pFeatureLayer.Valid = True Then ' or there will be an empty object ref

                pFeatureSelection = CType(pFeatureLayer, IFeatureSelection)
                pSelectionSet = pFeatureSelection.SelectionSet

                ' get the fields from the featureclass
                pFields = pFeatureLayer.FeatureClass.Fields
                j = 0

                ' 2020 - separate loops for polys and lines introduced (not most efficient)
                j = 0
                For j = 0 To lLineLayersFields.Count - 1
                    If lLineLayersFields(j).Layer = pFeatureLayer.Name Then

                        ' Get the Units of measure, if any
                        sUnit = lLineLayersFields(j).HabUnits
                        If sUnit = "Metres" Then
                            sUnit = "m"
                        ElseIf sUnit = "Kilometres" Then
                            sUnit = "km"
                        ElseIf sUnit = "Square Metres" Then
                            sUnit = "m^2"
                        ElseIf sUnit = "Feet" Then
                            sUnit = "ft"
                        ElseIf sUnit = "Miles" Then
                            sUnit = "mi"
                        ElseIf sUnit = "Square Miles" Then
                            sUnit = "mi^2"
                        ElseIf sUnit = "Hectares" Then
                            sUnit = "ha"
                        ElseIf sUnit = "Acres" Then
                            sUnit = "ac"
                        ElseIf sUnit = "Hectometres" Then
                            sUnit = "hm"
                        ElseIf sUnit = "Dekametres" Then
                            sUnit = "dm"
                        ElseIf sUnit = "Square Kilometres" Then
                            sUnit = "km^2"
                        ElseIf sUnit = "None" Then
                            sUnit = "none"
                        Else
                            sUnit = "n/a"
                        End If

                        'MsgBox("Debug 2020: Check for the habitat class field for line layer. Is it <none> or 'not set'?: " & lLineLayersFields(j).HabClsField)

                        Try
                            iLengthField = pFields.FindField(lLineLayersFields(j).LengthField)
                        Catch ex As Exception
                            MsgBox("Error finding the field in line layer for length. FIPEX code 971")
                            Exit Sub
                        End Try
                        ' 

                        ' if we find class field being used then use an intermediate object pHabStatisticsObject
                        ' then use a list of these objects (lHabStatsMatrix) to keep track of total habitat by class
                        iClassCheckTemp = pFields.FindField(lLineLayersFields(j).HabClsField)
                        If iClassCheckTemp <> -1 And lLineLayersFields(j).HabClsField <> "<None>" _
                            And lLineLayersFields(j).HabClsField <> "Not set" And lLineLayersFields(j).HabClsField <> "<none>" Then

                            ' Reset the stats objects
                            pDataStats = New DataStatistics
                            pHabStatisticsObject = New HabStatisticsObject(Nothing, Nothing)
                            ' Clear the statsMatrix
                            lHabStatsMatrix = New List(Of HabStatisticsObject)

                            If pFeatureSelection.SelectionSet.Count <> 0 Then

                                pSelectionSet.Search(Nothing, False, pCursor)

                                pHabStatisticsObject = New HabStatisticsObject(Nothing, Nothing)
                                With pHabStatisticsObject
                                    .UniqueHabClass = "Classes"
                                    .HabQuantity = Nothing '***
                                End With
                                lHabStatsMatrix.Add(pHabStatisticsObject)

                                pSelectionSet.Search(Nothing, False, pCursor) ' THIS LINE MAY BE REDUNDANT (SEE ABOVE)
                                pFeatureCursor = CType(pCursor, IFeatureCursor)

                                pFeature = pFeatureCursor.NextFeature                            ' For each selected feature
                                Do While Not pFeature Is Nothing

                                    ' 2020 - do not add if feature is to be excluded
                                    bExclude = False
                                    SharedSubs.exclusions2020(bExclude, pFeature, pFeatureLayer)

                                    If bExclude = False Then
                                        ' =========================================
                                        ' ============ HABITAT STATS ==============
                                        ' The habitat class field could be a number or a string
                                        ' so the variable used to hold it is an ambiguous object (variant)
                                        vFeatHbClsVl = pFeature.Value(pFields.FindField(lLineLayersFields(j).HabClsField))

                                        ' Loop through each unique habitat class again
                                        ' and check if it matches the class value of the feature
                                        k = 1
                                        bClassFound = False
                                        iStatsMatrixIndex = 0

                                        Try
                                            sClass = Convert.ToString(vFeatHbClsVl)
                                        Catch ex As Exception
                                            MsgBox("The Habitat Class found in the " & lLineLayersFields(j).Layer & " was not convertible" _
                                            & " to type 'string'.  " & ex.Message)
                                            sClass = "not set"
                                        End Try
                                        If sClass = "" Then
                                            sClass = "not set"
                                        End If

                                        vHabTemp = pFeature.Value(pFields.FindField(lLineLayersFields(j).HabQuanField))

                                        Try
                                            dHabLength = Convert.ToDouble(vHabTemp)
                                        Catch ex As Exception
                                            MsgBox("The Habitat Quantity found in the " & lLineLayersFields(j).Layer & " was not convertible" _
                                            & " to type 'double'.  Null values in field may be responsible. " & ex.Message)
                                            dHabLength = 0
                                        End Try

                                        classComparer = New FindStatsClassPredicate2020(sClass)
                                        ' use the layer and sink ID to get a refined list of habitat stats for 
                                        ' this sink, layer combo
                                        iStatsMatrixIndex = lHabStatsMatrix.FindIndex(AddressOf classComparer.CompareStatsClass)
                                        If iStatsMatrixIndex = -1 Then
                                            bClassFound = False
                                            pHabStatisticsObject = New HabStatisticsObject(sClass, dHabLength)
                                            lHabStatsMatrix.Add(pHabStatisticsObject)
                                        Else
                                            bClassFound = True
                                            lHabStatsMatrix(iStatsMatrixIndex).HabQuantity = lHabStatsMatrix(iStatsMatrixIndex).HabQuantity + dHabLength
                                        End If
                                        ' ============ END HABITAT STATS ==============
                                        ' =========================================

                                    End If

                                    ' ====================================================
                                    ' ============ LENGTH / DISTANCE STATS ===============
                                    ' 2020 exclusions don't apply to line length fields
                                    ' 2020 get distance / length field and quantity separetely from habitat (no classes for lengths)
                                    Try
                                        vTemp = pFeature.Value(iLengthField)
                                    Catch ex As Exception
                                        MsgBox("Problem converting value from length attribute in line layer to double / decimal. Please ensure values in attribute are of type double. FIPEX code 972.")
                                        vTemp = 0
                                    End Try
                                    Try
                                        dTempQuan = Convert.ToDouble(pFeature.Value(iLengthField))
                                    Catch ex As Exception
                                        dTempQuan = 0
                                        MsgBox("DCI table creation error. The length found in field in " & pFeatureLayer.Name & " was not convertible" _
                                            & " to type 'double'.  Null values in field may be responsible. FIPEX Code 973." & ex.Message)
                                    End Try
                                    dTotalLength = dTotalLength + dTempQuan
                                    ' ========== END LENGTH / DISTANCE STATS =============
                                    ' ====================================================

                                    pFeature = pFeatureCursor.NextFeature

                                Loop     ' next selected feature
                            End If ' There is a selection set


                            ' If there are items in the stats matrix
                            If lHabStatsMatrix.Count <> 0 Then
                                k = 1
                                ' For each unique value in the matrix
                                ' (always skip first row of matrix as it is the 'column headings')
                                For k = 1 To lHabStatsMatrix.Count - 1
                                    'If bDBF = True Then

                                    pHabStatsObject_2 = New StatisticsObject_2(Nothing, Nothing, Nothing, Nothing, _
                                                                               Nothing, Nothing, Nothing, Nothing, _
                                                                               Nothing, Nothing, Nothing, Nothing, _
                                                                               Nothing, Nothing, Nothing)
                                    With pHabStatsObject_2
                                        .Layer = pFeatureLayer.Name
                                        .LayerID = pFeatureLayer.FeatureClass.FeatureClassID
                                        .bID = ID
                                        .bEID = iEID
                                        .bType = sType
                                        .Sink = f_sOutID
                                        .SinkEID = f_siOutEID
                                        .Direction = sDirection2
                                        .LengthOrHabitat = "habitat"
                                        .HabitatDimension = "length"
                                        .TotalImmedPath = sHabTypeKeyword
                                        .UniqueClass = CStr(lHabStatsMatrix(k).UniqueHabClass)
                                        .ClassName = CStr(lLineLayersFields(j).HabClsField)
                                        .Quantity = lHabStatsMatrix(k).HabQuantity
                                        .Unit = sUnit
                                    End With
                                    lHabStatsList.Add(pHabStatsObject_2)

                                Next

                            Else ' If there are no statistics

                                pHabStatsObject_2 = New StatisticsObject_2(Nothing, Nothing, Nothing, Nothing, _
                                                                           Nothing, Nothing, Nothing, Nothing, _
                                                                           Nothing, Nothing, Nothing, Nothing, _
                                                                           Nothing, Nothing, Nothing)
                                With pHabStatsObject_2
                                    .Layer = pFeatureLayer.Name
                                    .LayerID = pFeatureLayer.FeatureClass.FeatureClassID
                                    .bID = ID
                                    .bEID = iEID
                                    .bType = sType
                                    .Sink = f_sOutID
                                    .SinkEID = f_siOutEID
                                    .Direction = sDirection2
                                    .LengthOrHabitat = "habitat"
                                    .HabitatDimension = "length"
                                    .TotalImmedPath = sHabTypeKeyword
                                    .UniqueClass = "not set"
                                    .ClassName = CStr(lLineLayersFields(j).HabClsField)
                                    .Quantity = 0.0
                                    .Unit = sUnit
                                End With
                                lHabStatsList.Add(pHabStatsObject_2)

                            End If ' There are items in the statsmatrix

                            pHabStatsObject_2 = New StatisticsObject_2(Nothing, Nothing, Nothing, Nothing, _
                                                                          Nothing, Nothing, Nothing, Nothing, _
                                                                          Nothing, Nothing, Nothing, Nothing, _
                                                                          Nothing, Nothing, Nothing)
                            With pHabStatsObject_2
                                .Layer = pFeatureLayer.Name
                                .LayerID = pFeatureLayer.FeatureClass.FeatureClassID
                                .bID = ID
                                .bEID = iEID
                                .bType = sType
                                .Sink = f_sOutID
                                .SinkEID = f_siOutEID
                                .Direction = sDirection2
                                .LengthOrHabitat = "length"
                                .HabitatDimension = "length"
                                .TotalImmedPath = sHabTypeKeyword
                                .UniqueClass = "not set"
                                .ClassName = CStr(lLineLayersFields(j).HabClsField)
                                .Quantity = dTotalLength
                                .Unit = sUnit
                            End With
                            lHabStatsList.Add(pHabStatsObject_2)


                        Else   ' if the habitat class case field is not found

                            dHabLength = 0

                            If pFeatureSelection.SelectionSet.Count <> 0 Then

                                pFeatureSelection.SelectionSet.Search(Nothing, False, pCursor)
                                pFeatureCursor = CType(pCursor, IFeatureCursor)
                                pFeature = pFeatureCursor.NextFeature

                                ' Get the summary field and add the value to the
                                ' total for habitat area.
                                ' ** ==> Multiple fields could be added here in a 'for' loop.
                                iFieldVal = pFeatureCursor.FindField(lLineLayersFields(j).HabQuanField)

                                ' For each selected feature
                                m = 1
                                Do While Not pFeature Is Nothing

                                    ' 2020 - do not add if feature is to be excluded
                                    bExclude = False
                                    SharedSubs.exclusions2020(bExclude, pFeature, pFeatureLayer)

                                    If bExclude = False Then
                                        Try
                                            vTemp = pFeature.Value(iFieldVal)
                                        Catch ex As Exception
                                            MsgBox("Could not convert quantity field found in " + lLineLayersFields(j).Layer.ToString + _
                                                   " was not convertible to type 'double'.  Null values in field may be responsible. " & ex.Message)
                                            vTemp = 0
                                        End Try

                                        Try
                                            dTempQuan = Convert.ToDouble(vTemp)
                                        Catch ex As Exception
                                            MsgBox("Could not convert the habitat quantity value found in the " + _
                                            lLineLayersFields(j).Layer.ToString + ". The values in the " + _
                                            lLineLayersFields(j).HabQuanField.ToString + " was not convertable to type 'double'." + _
                                            ex.Message)
                                            dTempQuan = 0
                                        End Try

                                        dHabLength = dHabLength + dTempQuan
                                    End If

                                    ' 2020 get distance / length field and quantity separetely from habitat (no classes for lengths)
                                    Try
                                        vTemp = pFeature.Value(iLengthField)
                                    Catch ex As Exception
                                        MsgBox("Problem converting value from length attribute in line layer to double / decimal. Please ensure values in attribute are of type double. FIPEX code 972.")
                                        vTemp = 0
                                    End Try
                                    Try
                                        dTempQuan = Convert.ToDouble(pFeature.Value(iLengthField))
                                    Catch ex As Exception
                                        dTempQuan = 0
                                        MsgBox("DCI table creation error. The length found in field in " & pFeatureLayer.Name & " was not convertible" _
                                            & " to type 'double'.  Null values in field may be responsible. FIPEX Code 973." & ex.Message)
                                    End Try
                                    dTotalLength = dTotalLength + dTempQuan

                                    pFeature = pFeatureCursor.NextFeature
                                Loop     ' selected feature
                            End If ' there are selected features

                            ' If DBF tables are to be output
                            'If bDBF = True Then

                            pHabStatsObject_2 = New StatisticsObject_2(Nothing, Nothing, Nothing, Nothing, _
                                                                       Nothing, Nothing, Nothing, Nothing, _
                                                                       Nothing, Nothing, Nothing, Nothing, _
                                                                       Nothing, Nothing, Nothing)
                            With pHabStatsObject_2
                                .Layer = pFeatureLayer.Name
                                .LayerID = pFeatureLayer.FeatureClass.FeatureClassID
                                .bID = ID
                                .bEID = iEID
                                .bType = sType
                                .Sink = f_sOutID
                                .SinkEID = f_siOutEID
                                .Direction = sDirection2
                                .LengthOrHabitat = "habitat"
                                .HabitatDimension = "length"
                                .TotalImmedPath = sHabTypeKeyword
                                .UniqueClass = "not set"
                                .ClassName = CStr(lLineLayersFields(j).HabClsField)
                                .Quantity = dHabLength
                                .Unit = sUnit
                            End With
                            lHabStatsList.Add(pHabStatsObject_2)

                            pHabStatsObject_2 = New StatisticsObject_2(Nothing, Nothing, Nothing, Nothing, _
                                                                        Nothing, Nothing, Nothing, Nothing, _
                                                                        Nothing, Nothing, Nothing, Nothing, _
                                                                        Nothing, Nothing, Nothing)
                            With pHabStatsObject_2
                                .Layer = pFeatureLayer.Name
                                .LayerID = pFeatureLayer.FeatureClass.FeatureClassID
                                .bID = ID
                                .bEID = iEID
                                .bType = sType
                                .Sink = f_sOutID
                                .SinkEID = f_siOutEID
                                .Direction = sDirection2
                                .LengthOrHabitat = "length"
                                .HabitatDimension = "length"
                                .TotalImmedPath = sHabTypeKeyword
                                .UniqueClass = "not set"
                                .ClassName = CStr(lLineLayersFields(j).HabClsField)
                                .Quantity = dTotalLength
                                .Unit = sUnit
                            End With
                            lHabStatsList.Add(pHabStatsObject_2)

                        End If ' found habitat class field in line layer

                        ' increment the loop counter for
                        iLoopCount = iLoopCount + 1

                    End If     ' feature layer matches hab class layer
                Next    ' line layer

                ' ##########################################################
                ' ##########################################################
                'Sep 2020 - separate loops for polys and lines.
                '       not efficient, but easiest and running short on time.
                j = 0
                For j = 0 To lPolyLayersFields.Count - 1
                    If lPolyLayersFields(j).Layer = pFeatureLayer.Name Then

                        ' Get the Units of measure, if any
                        sUnit = lPolyLayersFields(j).HabUnitField
                        If sUnit = "Metres" Then
                            sUnit = "m"
                        ElseIf sUnit = "Kilometres" Then
                            sUnit = "km"
                        ElseIf sUnit = "Square Metres" Then
                            sUnit = "m^2"
                        ElseIf sUnit = "Feet" Then
                            sUnit = "ft"
                        ElseIf sUnit = "Miles" Then
                            sUnit = "mi"
                        ElseIf sUnit = "Square Miles" Then
                            sUnit = "mi^2"
                        ElseIf sUnit = "Hectares" Then
                            sUnit = "ha"
                        ElseIf sUnit = "Acres" Then
                            sUnit = "ac"
                        ElseIf sUnit = "Hectometres" Then
                            sUnit = "hm"
                        ElseIf sUnit = "Dekametres" Then
                            sUnit = "dm"
                        ElseIf sUnit = "Square Kilometres" Then
                            sUnit = "km^2"
                        ElseIf sUnit = "None" Then
                            sUnit = "none"
                        Else
                            sUnit = "n/a"
                        End If

                        iClassCheckTemp = pFields.FindField(lPolyLayersFields(j).HabClsField)
                        'If pFields.FindField(lLayersFields(j).ClsField) <> -1 Then
                        If iClassCheckTemp <> -1 And lPolyLayersFields(j).HabClsField <> "<None>" _
                            And lPolyLayersFields(j).HabClsField <> "Not set" And lPolyLayersFields(j).HabClsField <> "<none>" Then

                            ' Reset the stats objects
                            pDataStats = New DataStatistics
                            pHabStatisticsObject = New HabStatisticsObject(Nothing, Nothing)
                            ' Clear the statsMatrix
                            lHabStatsMatrix = New List(Of HabStatisticsObject)

                            If pFeatureSelection.SelectionSet.Count <> 0 Then
                                pSelectionSet.Search(Nothing, False, pCursor)

                                pHabStatisticsObject = New HabStatisticsObject(Nothing, Nothing)
                                With pHabStatisticsObject
                                    .UniqueHabClass = "Classes"
                                    .HabQuantity = Nothing '***
                                End With
                                lHabStatsMatrix.Add(pHabStatisticsObject)

                                pSelectionSet.Search(Nothing, False, pCursor) ' THIS LINE MAY BE REDUNDANT (SEE ABOVE)
                                pFeatureCursor = CType(pCursor, IFeatureCursor)

                                pFeature = pFeatureCursor.NextFeature          ' For each selected feature
                                Do While Not pFeature Is Nothing

                                    ' 2020 - do not add if feature is to be excluded
                                    bExclude = False
                                    SharedSubs.exclusions2020(bExclude, pFeature, pFeatureLayer)

                                    'pFields = pFeature.Fields  '** removed because should be redundant

                                    If bExclude = False Then

                                        ' The habitat class field could be a number or a string
                                        ' so the variable used to hold it is an ambiguous object (variant)
                                        vFeatHbClsVl = pFeature.Value(pFields.FindField(lPolyLayersFields(j).HabClsField))

                                        ' Loop through each unique habitat class again
                                        ' and check if it matches the class value of the feature
                                        k = 1
                                        bClassFound = False
                                        iStatsMatrixIndex = 0

                                        Try
                                            sClass = Convert.ToString(vFeatHbClsVl)
                                        Catch ex As Exception
                                            MsgBox("The Habitat Class found in the " & lPolyLayersFields(j).Layer & " was not convertible" _
                                            & " to type 'string'.  " & ex.Message)
                                            sClass = "not set"
                                        End Try
                                        If sClass = "" Then
                                            sClass = "not set"
                                        End If

                                        vHabTemp = pFeature.Value(pFields.FindField(lPolyLayersFields(j).HabQuanField))

                                        Try
                                            dHabArea = Convert.ToDouble(vHabTemp)
                                        Catch ex As Exception
                                            MsgBox("The Habitat Quantity found in the " & lPolyLayersFields(j).Layer & " was not convertible" _
                                            & " to type 'double'.  Null values in field may be responsible. " & ex.Message)
                                            dHabArea = 0
                                        End Try

                                        classComparer = New FindStatsClassPredicate2020(sClass)
                                        ' use the layer and sink ID to get a refined list of habitat stats for 
                                        ' this sink, layer combo
                                        iStatsMatrixIndex = lHabStatsMatrix.FindIndex(AddressOf classComparer.CompareStatsClass)
                                        If iStatsMatrixIndex = -1 Then
                                            bClassFound = False
                                            pHabStatisticsObject = New HabStatisticsObject(sClass, dHabArea)
                                            lHabStatsMatrix.Add(pHabStatisticsObject)
                                        Else
                                            bClassFound = True
                                            lHabStatsMatrix(iStatsMatrixIndex).HabQuantity = lHabStatsMatrix(iStatsMatrixIndex).HabQuantity + dHabArea
                                        End If
                                    End If

                                    pFeature = pFeatureCursor.NextFeature

                                Loop     ' selected feature
                            End If ' There is a selection set

                            ' If there are items in the stats matrix
                            If lHabStatsMatrix.Count <> 0 Then
                                k = 1
                                ' For each unique value in the matrix
                                ' (always skip first row of matrix as it is the 'column headings')
                                For k = 1 To lHabStatsMatrix.Count - 1
                                    'If bDBF = True Then

                                    pHabStatsObject_2 = New StatisticsObject_2(Nothing, Nothing, Nothing, Nothing, _
                                                                               Nothing, Nothing, Nothing, Nothing, _
                                                                               Nothing, Nothing, Nothing, Nothing, _
                                                                               Nothing, Nothing, Nothing)
                                    With pHabStatsObject_2
                                        .Layer = pFeatureLayer.Name
                                        .LayerID = pFeatureLayer.FeatureClass.FeatureClassID
                                        .bID = ID
                                        .bEID = iEID
                                        .bType = sType
                                        .Sink = f_sOutID
                                        .SinkEID = f_siOutEID
                                        .Direction = sDirection2
                                        .LengthOrHabitat = "habitat"
                                        .HabitatDimension = "area"
                                        .TotalImmedPath = sHabTypeKeyword
                                        .UniqueClass = CStr(lHabStatsMatrix(k).UniqueHabClass)
                                        .ClassName = CStr(lPolyLayersFields(j).HabClsField)
                                        .Quantity = lHabStatsMatrix(k).HabQuantity
                                        .Unit = sUnit
                                    End With
                                    lHabStatsList.Add(pHabStatsObject_2)
                                Next
                            Else ' If there are no statistics

                                pHabStatsObject_2 = New StatisticsObject_2(Nothing, Nothing, Nothing, Nothing, _
                                                                           Nothing, Nothing, Nothing, Nothing, _
                                                                           Nothing, Nothing, Nothing, Nothing, _
                                                                           Nothing, Nothing, Nothing)
                                With pHabStatsObject_2
                                    .Layer = pFeatureLayer.Name
                                    .LayerID = pFeatureLayer.FeatureClass.FeatureClassID
                                    .bID = ID
                                    .bEID = iEID
                                    .bType = sType
                                    .Sink = f_sOutID
                                    .SinkEID = f_siOutEID
                                    .Direction = sDirection2
                                    .LengthOrHabitat = "habitat"
                                    .HabitatDimension = "area"
                                    .TotalImmedPath = sHabTypeKeyword
                                    .UniqueClass = "not set"
                                    .ClassName = CStr(lPolyLayersFields(j).HabClsField)
                                    .Quantity = 0.0
                                    .Unit = sUnit
                                End With
                                lHabStatsList.Add(pHabStatsObject_2)

                            End If ' There are items in the statsmatrix
                        Else   ' if the habitat class case field is not found

                            dHabArea = 0

                            If pFeatureSelection.SelectionSet.Count <> 0 Then

                                pFeatureSelection.SelectionSet.Search(Nothing, False, pCursor)
                                pFeatureCursor = CType(pCursor, IFeatureCursor)
                                pFeature = pFeatureCursor.NextFeature

                                ' Get the summary field and add the value to the total for habitat area.
                                ' ** ==> Multiple fields could be added here in a 'for' loop.
                                iFieldVal = pFeatureCursor.FindField(lPolyLayersFields(j).HabQuanField)

                                ' For each selected feature
                                m = 1
                                Do While Not pFeature Is Nothing

                                    ' 2020 - do not add if feature is to be excluded
                                    bExclude = False
                                    SharedSubs.exclusions2020(bExclude, pFeature, pFeatureLayer)

                                    If bExclude = False Then

                                        Try
                                            vTemp = pFeature.Value(iFieldVal)
                                        Catch ex As Exception
                                            MsgBox("Could not convert quantity field found in " + lPolyLayersFields(j).Layer.ToString + _
                                                   " was not convertible to type 'double'.  Null values in field may be responsible. " & ex.Message)
                                            vTemp = 0
                                        End Try
                                        Try
                                            dTempQuan = Convert.ToDouble(vTemp)
                                        Catch ex As Exception
                                            MsgBox("Could not convert the habitat quantity value found in the " + _
                                            lPolyLayersFields(j).Layer.ToString + "layer. The " + _
                                            +lPolyLayersFields(j).HabQuanField.ToString + " field was not convertable to type 'double'." + _
                                            ex.Message)

                                            dTempQuan = 0
                                        End Try
                                        ' Insert into the corresponding column of the second
                                        ' row the updated habitat area measurement.
                                        dHabArea = dHabArea + dTempQuan
                                    End If

                                    pFeature = pFeatureCursor.NextFeature
                                Loop     ' selected feature
                            End If ' there are selected features

                            ' If DBF tables are to be output
                            'If bDBF = True Then

                            pHabStatsObject_2 = New StatisticsObject_2(Nothing, Nothing, Nothing, Nothing, _
                                                                       Nothing, Nothing, Nothing, Nothing, _
                                                                       Nothing, Nothing, Nothing, Nothing, _
                                                                       Nothing, Nothing, Nothing)
                            With pHabStatsObject_2
                                .Layer = pFeatureLayer.Name
                                .LayerID = pFeatureLayer.FeatureClass.FeatureClassID
                                .bID = ID
                                .bEID = iEID
                                .bType = sType
                                .Sink = f_sOutID
                                .SinkEID = f_siOutEID
                                .Direction = sDirection2
                                .LengthOrHabitat = "habitat"
                                .HabitatDimension = "area"
                                .TotalImmedPath = sHabTypeKeyword
                                .UniqueClass = "not set"
                                .ClassName = CStr(lPolyLayersFields(j).HabClsField)
                                .Quantity = dHabArea
                                .Unit = sUnit
                            End With
                            lHabStatsList.Add(pHabStatsObject_2)

                        End If ' found habitat class field in layer

                        ' increment the loop counter for
                        iLoopCount = iLoopCount + 1

                    End If  ' feature layer matches hab class layer
                Next    ' poly layer
            End If ' featurelayer is valid
            pFeatureLayer = CType(pEnumLayer.Next(), IFeatureLayer)
        Loop

    End Sub

    Public Class FindStatsClassPredicate2020
        ' this class should help return a double-check 
        ' list object of Statistics where the layer matches 
        ' and the sink/barr EID matches as well.  
        ' tutorial here: http://social.msdn.microsoft.com/Forums/en-US/vbgeneral/thread/bad5193a-3bf1-4675-8a05-625f46f9c158/
        Private _Class As String

        Public Sub New(ByVal class2 As String)
            Me._Class = class2
        End Sub

        ' this is a double-check and should 
        ' return true if there's a match between
        ' BOTH the layer and EID of the sink/barrier
        Public Function CompareStatsClass(ByVal obj As HabStatisticsObject) As Boolean
            Return (_Class = obj.UniqueHabClass)
        End Function
    End Class

    Public Shared Function flagcheck2021(ByRef pBarriersList As IEnumNetEID, ByRef pEdgeFlags As IEnumNetEID, ByRef pJuncFlags As IEnumNetEID) As String
        ' Function:     Flag Check
        ' Author:       Greig Oldford
        ' Date Created: March 26th, 2008
        ' Date Translated: May 11 - 19, 2009 
        ' Description:  This function will check to see if initially the flags are on barriers
        '               or not. It should stop the analysis if some flags are on barriers and some
        '               are not - it should be one or the other.
        '               It will return one of three strings - "barrier", "nonbarr", or "error"
        '
        ' Notes:   March 27, 2008 -> Not currently checking for edge flags since there is no support for
        '                      them yet.
        '          Apr 28, 2021 -> Copied this function here from Analysis.vb, made public
        Dim i As Integer
        Dim m As Integer
        Dim pFlagsOnBarrGEN As IEnumNetEIDBuilderGEN  ' list holds flags on barriers
        Dim pFlagsNoBarrGEN As IEnumNetEIDBuilderGEN  ' list holds flag not on barriers
        pFlagsOnBarrGEN = New EnumNetEIDArray
        pFlagsNoBarrGEN = New EnumNetEIDArray

        Dim pFlagsOnBarr As IEnumNetEID
        Dim pFlagsNoBarr As IEnumNetEID
        Dim flagBarrier As Boolean
        Dim iEID As Integer

        pJuncFlags.Reset()
        i = 0

        ' For each flag
        For i = 0 To pJuncFlags.Count - 1

            flagBarrier = False     ' assume flag is not on barrier
            iEID = pJuncFlags.Next()  ' get the EID of flag
            m = 0
            pBarriersList.Reset()

            ' For each barrier
            For m = 0 To pBarriersList.Count - 1
                'If endEID = pOriginalBarriersList(m) Then 'VB.NET
                If iEID = pBarriersList.Next() Then
                    flagBarrier = True
                End If
            Next

            If flagBarrier = True Then  'put EID in flags on barrier list

                ' THIS LIST COULD BE USED IN FUTURE TO FILTER BAD FLAGS OUT
                ' I.E. - check which flags are on barriers and remove only those ones automatically.
                pFlagsOnBarrGEN.Add(iEID)

            Else   ' put EID in flags not on barrier list
                pFlagsNoBarrGEN.Add(iEID)
            End If
        Next

        ' QI to get "next" and "count"
        pFlagsOnBarr = CType(pFlagsOnBarrGEN, IEnumNetEID)
        pFlagsNoBarr = CType(pFlagsNoBarrGEN, IEnumNetEID)

        If pFlagsOnBarr.Count = pJuncFlags.Count Then
            flagcheck2021 = "barriers"
            'return "barriers"? ' should be a return in VB.Net I think... but this works
        ElseIf pFlagsNoBarr.Count = pJuncFlags.Count Then
            flagcheck2021 = "nonbarr"
        Else
            MsgBox("Inconsistent flag placement." + vbCrLf + _
            "Barrier flags: " & pFlagsOnBarr.Count & vbCrLf & _
            " Non-barrier flags: " & pFlagsNoBarr.Count)
            flagcheck2021 = "error"
        End If
    End Function
    Public Shared Function TraceFlowSolverSetup() As ITraceFlowSolver
        ' Prepares the network for tracing
        ' duplicated in 2021 from private function in analysis.vb

        Dim pNetworkAnalysisExt As INetworkAnalysisExt
        Dim pGeometricNetwork As IGeometricNetwork
        Dim pUtilityNetwork As IUtilityNetwork
        Dim pNetSolver As INetSolver
        Dim pNetworkAnalysisExtBarriers As INetworkAnalysisExtBarriers
        Dim pNetworkAnalysisExtFlags As INetworkAnalysisExtFlags
        Dim pEdgeElementBarriers As INetElementBarriers     ' Should these next two be INetElementBarriersGEN?
        Dim pJunctionElementBarriers As INetElementBarriers
        Dim pSelectionSetBarriers As ISelectionSetBarriers
        Dim pFeatureLayer As IFeatureLayer
        Dim pFlagDisplay As IFlagDisplay
        Dim pEdgeFlagDisplay As IEdgeFlagDisplay
        Dim pEdgeFlags() As IEdgeFlag
        Dim pJunctionFlags() As IJunctionFlag
        Dim pNetFlag As INetFlag
        Dim pEdgeFlag As IEdgeFlag
        'Dim pTraceFlowSolver As ITraceFlowSolver
        Dim pTraceTasks As ITraceTasks
        Dim pNetworkAnalysisExtWeightFilter As INetworkAnalysisExtWeightFilter
        Dim pNetSchema As INetSchema
        Dim pNetWeight As INetWeight
        Dim eWeightFilterType As esriWeightFilterType
        Dim pNetSolverWeights As INetSolverWeights
        Dim pTraceFlowSolver As ITraceFlowSolver

        ' Must change from Object type in VB.Net to Variant in VB6
        Dim lngFromValues() As Object
        Dim lngToValues() As Object

        Dim lngFeatureLayerCount As Integer
        Dim lngEdgeFlagCount As Integer
        Dim lngJunctionFlagCount As Integer
        Dim binFeatureLayerDisabled As Boolean
        Dim binApplyNotOperator As Boolean
        Dim lngFilterRangeCount As Integer
        Dim i As Integer

        ' copied these module level vars just as a test - GO May 2021
        ' should be reverted to local vars
        Dim m_FiPEx__1 As FishPassageExtension
        Dim m_UNAExt As ESRI.ArcGIS.EditorExt.IUtilityNetworkAnalysisExt
        Dim m_pNetworkAnalysisExt As INetworkAnalysisExt

        If m_FiPEx__1 Is Nothing Then
            m_FiPEx__1 = FiPEX_ArcMap_10p4_up_AddIn_dotNet45_2020.FishPassageExtension.GetExtension()
        End If
        If m_UNAExt Is Nothing Then
            m_UNAExt = FiPEX_ArcMap_10p4_up_AddIn_dotNet45_2020.FishPassageExtension.GetUNAExt
        End If
        'Dim FiPEx__1 As FishPassageExtension = FiPEX_AddIn_dotNet35_2.FishPassageExtension.GetExtension
        'Dim pUNAExt As IUtilityNetworkAnalysisExt = FiPEX_AddIn_dotNet35_2.FishPassageExtension.GetUNAExt
        If m_pNetworkAnalysisExt Is Nothing Then
            m_pNetworkAnalysisExt = CType(m_UNAExt, INetworkAnalysisExt)
        End If

        ' Get reference to the current network through Utility Network interface
        pNetworkAnalysisExt = CType(m_UNAExt, INetworkAnalysisExt)

        ' Assign a reference to the current geometric network via a local variable
        pGeometricNetwork = pNetworkAnalysisExt.CurrentNetwork

        ' Assign the network to local IUtilityNetwork variable
        ' however, do it through GeometricNetwork which returns
        ' a QI for utilitynetwork using INetwork.... ?

        ' Not all members of IUtilityNetwork be directly calleable
        ' With VB.NET and it has been recommended to use IUtilityNetworkGEN
        ' instead...

        pUtilityNetwork = CType(pGeometricNetwork.Network, IUtilityNetwork)

        ' initialize the trace flow solver and set the network
        pNetSolver = New TraceFlowSolver
        pNetSolver.SourceNetwork = pUtilityNetwork

        ' Get barriers for the network
        ' QI for the interface using the IUtilityNetworkAnalysisExt
        pNetworkAnalysisExtBarriers = CType(m_UNAExt, INetworkAnalysisExtBarriers)
        ' Get the element barriers
        pNetworkAnalysisExtBarriers.CreateElementBarriers(pJunctionElementBarriers, pEdgeElementBarriers)
        ' Get the selection set barriers
        pNetworkAnalysisExtBarriers.CreateSelectionBarriers(pSelectionSetBarriers)
        ' set the barriers for the network solver
        pNetSolver.ElementBarriers(esriElementType.esriETEdge) = pEdgeElementBarriers
        pNetSolver.ElementBarriers(esriElementType.esriETJunction) = pJunctionElementBarriers
        pNetSolver.SelectionSetBarriers = pSelectionSetBarriers


        ' set the barriers for the network solver
        ' for each feature determine if it is enabled.
        ' if it is disabled notify network solver
        ' determine the number of feature layers belonging to network
        lngFeatureLayerCount = pNetworkAnalysisExt.FeatureLayerCount
        For i = 0 To lngFeatureLayerCount - 1
            ' get the next feature layer and determine if it is disabled
            ' assign the feature layer to a local IFeatureLayer variable
            pFeatureLayer = pNetworkAnalysisExt.FeatureLayer(i)
            ' determine if the feature layer is disabled
            binFeatureLayerDisabled = pNetworkAnalysisExtBarriers.GetDisabledLayer(pFeatureLayer)
            ' if it is disabled then notify the network solver
            If binFeatureLayerDisabled Then
                pNetSolver.DisableElementClass(pFeatureLayer.FeatureClass.FeatureClassID)
            End If
        Next

        ' set up the weight filters for the network
        ' QI for the INetworkAnalysisEctWeightFilter interface using
        ' the UtilityNetworkAnalysisExt
        pNetworkAnalysisExtWeightFilter = CType(m_UNAExt, INetworkAnalysisExtWeightFilter)
        ' QI for the INetSolverWeights interace using the INetSolver
        pNetSolverWeights = CType(pNetSolver, INetSolverWeights)
        ' QI for the INetSchema interface using IUtilityNetwork
        pNetSchema = CType(pUtilityNetwork, INetSchema)

        With pNetworkAnalysisExtWeightFilter

            ' Create the junction weight filter
            lngFilterRangeCount = pNetworkAnalysisExtWeightFilter.FilterRangeCount(esriElementType.esriETJunction)
            ' If there are any weight filters
            If lngFilterRangeCount > 0 Then
                ' get a NetWeight object from the INetSchema interface
                pNetWeight = pNetSchema.WeightByName(.JunctionWeightFilterName)
                ' get the type and Not operator status from the InetworkAnalysisExtWeightFileter interface
                .GetFilterType(esriElementType.esriETJunction, eWeightFilterType, binApplyNotOperator)
                ' redimension the weight filter ranges arrays and get the ranges
                ReDim lngFromValues(0 To lngFilterRangeCount - 1)
                ReDim lngToValues(0 To lngFilterRangeCount - 1)
                For i = 0 To lngFilterRangeCount - 1
                    .GetFilterRange(esriElementType.esriETJunction, i, lngFromValues(i), lngToValues(i))
                Next
                ' add the filter ranges to the network solver
                pNetSolverWeights.SetFilterRanges(esriElementType.esriETJunction, lngFilterRangeCount, lngFromValues(0), lngToValues(0))

            End If


            ' create the edge weight filters
            lngFilterRangeCount = pNetworkAnalysisExtWeightFilter.FilterRangeCount(esriElementType.esriETEdge)
            If lngFilterRangeCount > 0 Then

                ' get the type and Not operator status from the INetworkAnalysisExtWeightFilter interface
                .GetFilterType(esriElementType.esriETEdge, eWeightFilterType, binApplyNotOperator)

                ' get a NetWeight object from the INetSchema interface
                pNetWeight = pNetSchema.WeightByName(.FromToEdgeWeightFilterName)
                ' add the weight filter to the network solver
                pNetSolverWeights.FromToEdgeFilterWeight = pNetWeight

                ' get a NetWeight object from the INetScema interface
                pNetWeight = pNetSchema.WeightByName(.ToFromEdgeWeightFilterName)
                ' add the weight filter to the network solver
                pNetSolverWeights.ToFromEdgeFilterWeight = pNetWeight

                'get the filter ranges and apply them to the network solver
                ReDim lngFromValues(0 To lngFilterRangeCount - 1)
                ReDim lngToValues(0 To lngFilterRangeCount - 1)
                For i = 0 To lngFilterRangeCount - 1
                    .GetFilterRange(esriElementType.esriETEdge, i, lngFromValues(i), lngToValues(i))
                Next

                pNetSolverWeights.SetFilterType(esriElementType.esriETEdge, eWeightFilterType, binApplyNotOperator)
                pNetSolverWeights.SetFilterRanges(esriElementType.esriETEdge, lngFilterRangeCount, lngFromValues(0), lngToValues(0))

            End If

        End With


        ' assign the flags to the Network Solver
        ' get the edge flags
        ' QI for the ITraceFlowSolver interface using INetSolver Interface
        pTraceFlowSolver = CType(pNetSolver, ITraceFlowSolver)
        ' QI for the INetworkAnalysisExtFlags interface using the IUtilitNetworkAnalysisExt
        pNetworkAnalysisExtFlags = CType(m_UNAExt, INetworkAnalysisExtFlags)
        ' determine the number of edge flags on the current network
        lngEdgeFlagCount = pNetworkAnalysisExtFlags.EdgeFlagCount

        ' if there are edge flags then
        If Not lngEdgeFlagCount = 0 Then
            ' Redimension the array to hold the correct number of edge flags
            ReDim pEdgeFlags(0 To lngEdgeFlagCount - 1)
            For i = 0 To lngEdgeFlagCount - 1
                ' assign a local variable for IFlagDisplay and IEdgeFlagDisplay variables
                pFlagDisplay = CType(pNetworkAnalysisExtFlags.EdgeFlag(i), IFlagDisplay)
                pEdgeFlagDisplay = CType(pFlagDisplay, IEdgeFlagDisplay)

                ' co-create(?)a new EdgeFlag Object
                pNetFlag = New EdgeFlag
                pEdgeFlag = CType(pNetFlag, IEdgeFlag)
                ' assign the properties of the EdgeFlagDisplay object to he EdgeFlag Object
                ' I think this is where you could determine the distance along the line
                ' an edge flag is. 
                pEdgeFlag.Position = Convert.ToSingle(pEdgeFlagDisplay.Percentage)
                pNetFlag.UserClassID = pFlagDisplay.FeatureClassID
                pNetFlag.UserID = pFlagDisplay.FID
                pNetFlag.UserSubID = pFlagDisplay.SubID
                ' add the new EdgeFlag object to the array
                pEdgeFlags(i) = CType(pNetFlag, IEdgeFlag)
            Next
            ' add the edge flags to the network solver
            'pTraceFlowSolver.PutEdgeOrigins(lngEdgeFlagCount, pEdgeFlags(0))
        End If


        ' Get the junction flags
        ' determine the number of junction flags on the network
        lngJunctionFlagCount = pNetworkAnalysisExtFlags.JunctionFlagCount
        ' only execute this if there are juntion flags
        If Not lngJunctionFlagCount = 0 Then
            ' redimension the array to hold the correct number of junction flags
            ReDim pJunctionFlags(0 To lngJunctionFlagCount - 1)
            For i = 0 To lngJunctionFlagCount - 1

                ' assign to a local IFlagDisplay variable
                pFlagDisplay = CType(pNetworkAnalysisExtFlags.JunctionFlag(i), IFlagDisplay)
                ' co-create a new JunctionFlag object
                pNetFlag = New JunctionFlag
                ' assign the properties of the JunctionFlagDisplay object to the JunctionFlag object
                pNetFlag.UserClassID = pFlagDisplay.FeatureClassID
                pNetFlag.UserID = pFlagDisplay.FID
                pNetFlag.UserSubID = pFlagDisplay.SubID
                ' add the new junction flag to the array of junction flags
                pJunctionFlags(i) = CType(pNetFlag, IJunctionFlag)
            Next
            'add the junction flags to the network solver
            ' This function is not calleable from VB.NET:
            'pTraceFlowSolver.PutJunctionOrigins(lngJunctionFlagCount, pJunctionFlags(0))
            ' So we use this function
            Dim pTraceFlowSolverGEN As ITraceFlowSolverGEN
            pTraceFlowSolverGEN = CType(pNetSolver, ITraceFlowSolverGEN)
            pTraceFlowSolverGEN.PutJunctionOrigins(pJunctionFlags)

        End If

        ' set the option for tracing on indeterminate flow
        ' QI for the ITraceTasksinterface using IUtilityNetworkAnlysisExt
        pTraceTasks = CType(m_UNAExt, ITraceTasks)
        'pTraceFlowSolver.TraceIndeterminateFlow = pTraceTasks.TraceIndeterminateFlow
        pTraceFlowSolver.TraceIndeterminateFlow = True
        pTraceTasks.TraceIndeterminateFlow = True

        ' pass the traceFlowSolver object back to the network solver
        TraceFlowSolverSetup = pTraceFlowSolver

    End Function

    Public Shared Sub ResetFlagsBarriers(ByRef pOriginalBarriersList As IEnumNetEID, ByRef pOriginalEdgeFlagsList As IEnumNetEID,
                                  ByRef pOriginaljuncFlagsList As IEnumNetEID, ByRef pUtilityNetworkAnalysisExt As IUtilityNetworkAnalysisExt)

        ' created 2020 by G Oldford
        ' duplicated 2021 - duplicate of private sub in analysis.vb

        Dim bEID, iEID, bFCID, bFID, bSubID, iFID, iFCID, iSubID, m As Integer
        Dim pFlagDisplay As IFlagDisplay
        Dim pSymbol As ISymbol
        Dim pJuncFlagDisplay As IJunctionFlagDisplay
        Dim pEdgeFlagDisplay As IEdgeFlagDisplay
        Dim pNetworkAnalysisExtBarriers As INetworkAnalysisExtBarriers
        Dim pNetworkAnalysisExtFlags As INetworkAnalysisExtFlags
        Dim pFlagSymbol, pBarrierSymbol As ISimpleMarkerSymbol

        GetSymbols(pFlagSymbol, pBarrierSymbol)

        Dim pNetworkAnalysisExt As INetworkAnalysisExt
        Dim pNetwork As INetwork
        Dim pNetElements As INetElements
        pNetworkAnalysisExt = CType(pUtilityNetworkAnalysisExt, INetworkAnalysisExt)
        pNetworkAnalysisExtFlags = CType(pUtilityNetworkAnalysisExt, INetworkAnalysisExtFlags)
        pNetworkAnalysisExtBarriers = CType(pUtilityNetworkAnalysisExt, INetworkAnalysisExtBarriers)
        Dim pGeometricNetwork As IGeometricNetwork = pNetworkAnalysisExt.CurrentNetwork
        pGeometricNetwork = pNetworkAnalysisExt.CurrentNetwork
        pNetwork = pGeometricNetwork.Network
        pNetElements = CType(pNetwork, INetElements)

        pNetworkAnalysisExtBarriers.ClearBarriers()

        ' ========================= RESET BARRIERS ===========================
        '               Reset things the way the user had them

        m = 0
        pOriginalBarriersList.Reset()
        For m = 0 To pOriginalBarriersList.Count - 1
            bEID = pOriginalBarriersList.Next()
            pNetElements.QueryIDs(bEID, esriElementType.esriETJunction, bFCID, bFID, bSubID)

            ' Display the barriers as a JunctionFlagDisplay type
            pFlagDisplay = New JunctionFlagDisplay
            pSymbol = CType(pBarrierSymbol, ISymbol)
            With pFlagDisplay
                .FeatureClassID = bFCID
                .FID = bFID
                .Geometry = pGeometricNetwork.GeometryForJunctionEID(bEID)
                .Symbol = pSymbol
            End With

            ' Add the flags to the logical network
            pJuncFlagDisplay = CType(pFlagDisplay, IJunctionFlagDisplay)
            pNetworkAnalysisExtBarriers.AddJunctionBarrier(pJuncFlagDisplay)
        Next
        ' ====================== END RESET BARRIERS ===========================
        'MsgBox("Debug:62")
        ' Clear current flags
        pNetworkAnalysisExtFlags.ClearFlags()

        ' ======================== RESET FLAGS ================================
        ' restore all EDGE flags
        m = 0
        pOriginalEdgeFlagsList.Reset()
        For m = 0 To pOriginalEdgeFlagsList.Count - 1

            iEID = pOriginalEdgeFlagsList.Next()
            ' Query the corresponding user ID's to the element ID
            pNetElements.QueryIDs(iEID, esriElementType.esriETEdge, iFCID, iFID, iSubID)

            ' Display the flags as a JunfctionFlagDisplay type
            pFlagDisplay = New EdgeFlagDisplay
            pSymbol = CType(pFlagSymbol, ISymbol)
            With pFlagDisplay
                .FeatureClassID = iFCID
                .FID = iFID
                .Geometry = pGeometricNetwork.GeometryForEdgeEID(iEID)
                .Symbol = pSymbol
            End With

            ' Add the flags to the logical network
            pEdgeFlagDisplay = CType(pFlagDisplay, IEdgeFlagDisplay)
            pNetworkAnalysisExtFlags.AddEdgeFlag(pEdgeFlagDisplay)
        Next

        ' restore all JUNCTION Flags
        m = 0
        pOriginaljuncFlagsList.Reset()
        For m = 0 To pOriginaljuncFlagsList.Count - 1

            iEID = pOriginaljuncFlagsList.Next()
            ' Query the corresponding user ID's to the element ID
            pNetElements.QueryIDs(iEID, esriElementType.esriETJunction, iFCID, iFID, iSubID)

            ' Display the flags as a JunfctionFlagDisplay type
            pFlagDisplay = New JunctionFlagDisplay
            pSymbol = CType(pFlagSymbol, ISymbol)
            With pFlagDisplay
                .FeatureClassID = iFCID
                .FID = iFID
                .Geometry = pGeometricNetwork.GeometryForJunctionEID(iEID)
                .Symbol = pSymbol
            End With

            ' Add the flags to the logical network
            pJuncFlagDisplay = CType(pFlagDisplay, IJunctionFlagDisplay)
            pNetworkAnalysisExtFlags.AddJunctionFlag(pJuncFlagDisplay)

        Next


        ' =========================== END RESET FLAGS =====================
    End Sub

    Public Shared Function GetBarrierID(ByVal iFCID As Integer, ByVal iFID As Integer, ByVal lBarrierIDs As List(Of BarrierIDObj)) As IDandType
        ' copied from analysis.vb May 2021 - GO

        ' =============== FLAG ON POINT WITH BNUMBER? =========================
        'returns user-set ID or Object ID
        ' This section checks whether there is a BNumber Field
        ' In which case it will use this field to identify the 
        ' barrier in the output.
        ' 1.0 For each layer in the table of contents
        '   2.0 If the layer is a FeatureLayer
        '     3.0 If the layer ClassID matches the layer of the CURRENT FLAG
        '     3.1 Get the field values for the feature
        '       4.0 For each of the Barrier IDs in the list
        '         5.0 If the layer in the BarrierID list matches the layer in the TOC
        '         5.1 Get the name of the field
        '           6.0 If the field has something in it
        '             7.0 Set the ID of the flag (sOID) equal to that value
        '
        Dim bCheck As Boolean
        Dim j, k As Integer

        Dim pDoc As IDocument = My.ArcMap.Application.Document
        Dim pMxDoc As IMxDocument = CType(pDoc, IMxDocument)
        Dim pMap As IMap = pMxDoc.FocusMap
        Dim pFLayer As IFeatureLayer
        Dim pFeatureClass As IFeatureClass
        Dim pFeature As IFeature
        Dim pFields As IFields
        Dim sBarrierIDField As String
        Dim iBarrierIds, iOID As Integer
        Dim sOutID As String
        Dim q As Integer
        Dim sAncillaryRole As String

        bCheck = False
        For j = 0 To pMap.LayerCount - 1
            If pMap.Layer(j).Valid = True Then
                If TypeOf pMap.Layer(j) Is IFeatureLayer Then
                    pFLayer = CType(pMap.Layer(j), IFeatureLayer)
                    If pFLayer.FeatureClass.FeatureClassID = iFCID Then

                        pFeatureClass = pFLayer.FeatureClass
                        pFeature = pFeatureClass.GetFeature(iFID)
                        pFields = pFeature.Fields
                        iBarrierIds = lBarrierIDs.Count

                        For k = 0 To iBarrierIds - 1
                            If lBarrierIDs.Item(k).Layer = pFLayer.Name Then
                                sBarrierIDField = lBarrierIDs.Item(k).Field
                                If pFields.FindField(sBarrierIDField) <> -1 Then
                                    Try
                                        sOutID = Convert.ToString(pFeature.Value(pFields.FindField(sBarrierIDField)))
                                    Catch ex As Exception
                                        MsgBox("Could not convert barrier ID to string.")
                                        sOutID = "Unknown"
                                    End Try
                                    bCheck = True
                                End If ' names match
                            End If
                        Next ' barrierIDField
                        If bCheck = False Then ' names don't match - set to default field
                            'sOutID = Convert.ToString(pFeature.Value(pFields.FindField("OBJECTID")))
                            ' If the layer has an ancillaryRole, check to see if the feature
                            q = pFields.FindField("AncillaryRole")
                            If Not q = -1 Then
                                sAncillaryRole = Convert.ToString(pFeature.Value(q))
                                If sAncillaryRole = "2" Then
                                    sOutID = "Sink"
                                    iOID = pFeature.OID.ToString ' Added Feb 18, 2012
                                End If
                            Else
                                sOutID = pFeature.OID.ToString
                            End If
                        End If
                    End If ' feature class in map equals feature class of flag
                End If
            End If
        Next

        Dim pIDAndType As New IDandType(Nothing, Nothing)

        If sOutID = "Sink" Then
            pIDAndType.BarrID = iOID.ToString
            pIDAndType.BarrIDType = sOutID
        Else
            pIDAndType.BarrID = sOutID
            pIDAndType.BarrIDType = "Barrier"
        End If

        GetBarrierID = pIDAndType

    End Function

    Public Shared Sub GetSymbols(ByRef pFlagSymbol As ISimpleMarkerSymbol, ByRef pBarrierSymbol As ISimpleMarkerSymbol)

        ' =========== GET SYMBOLS ============

        Dim pSimpleMarkerSymbol As ISimpleMarkerSymbol = New SimpleMarkerSymbol
        Dim pRgbColor As IRgbColor = New RgbColor

        ' For the Flag marker
        With pRgbColor
            .Red = 0
            .Green = 255
            .Blue = 0
        End With
        pSimpleMarkerSymbol = New SimpleMarkerSymbol
        With pSimpleMarkerSymbol
            .Color = pRgbColor
            .Style = esriSimpleMarkerStyle.esriSMSSquare
            .Outline = True
            .Size = 10
        End With

        ' Result is a global variable containing a flag marker
        pFlagSymbol = pSimpleMarkerSymbol
        pRgbColor = New RgbColor

        ' Set the barrier symbol color and parameters
        With pRgbColor
            .Red = 255
            .Green = 0
            .Blue = 0
        End With

        pSimpleMarkerSymbol = New SimpleMarkerSymbol
        With pSimpleMarkerSymbol
            .Color = pRgbColor
            .Style = esriSimpleMarkerStyle.esriSMSX
            .Outline = True
            .Size = 10
        End With

        ' Result is a global variable containing a barrier marker
        pBarrierSymbol = pSimpleMarkerSymbol

        ' ============= END GET SYMBOLS ==============
    End Sub

    Public Shared Sub ReturnSelectionSetList(ByRef gEID As Integer, ByRef sLabelField As String, _
                                             ByRef lSelectAndUpdateFeaturesObject As List(Of SelectAndUpdateFeaturesObject),
                                             ByRef pNetworkAnalysisExtResults As INetworkAnalysisExtResults,
                                             ByRef pNetworkAnalysisExt As INetworkAnalysisExt,
                                             ByRef pNodesSet As IEnumNetEID,
                                             ByRef pEdgesSet As IEnumNetEID)

        Dim pEnumLayer As IEnumLayer

        Dim pUID As New UID
        Dim pID As New UID
        Dim pMap As IMap
        Dim pMxDoc As IMxDocument
        Dim pFLyrSlct As IFeatureLayer
        Dim fEID, k As Integer
        Dim pDoc As IDocument = My.ArcMap.Application.Document
        pMxDoc = CType(pDoc, IMxDocument)
        pMap = pMxDoc.FocusMap
        pMap.ClearSelection()

        ' Make sure all the layers in the TOC are selectable
        For i = 0 To pMap.LayerCount - 1
            If pMap.Layer(i).Valid = True Then
                If TypeOf pMap.Layer(i) Is IFeatureLayer Then
                    pFLyrSlct = CType(pMap.Layer(i), IFeatureLayer)
                    pFLyrSlct.Selectable = True
                End If
            End If
        Next

        pNetworkAnalysisExtResults.ClearResults()
        pNetworkAnalysisExtResults.CreateSelection(pNodesSet, pEdgesSet)

        ' enumerate layers in TOC
        Try
            pUID.Value = "{E156D7E5-22AF-11D3-9F99-00C04F6BC78E}"
            pEnumLayer = pMap.Layers(pUID, True)
            If pEnumLayer Is Nothing Then
                MsgBox("Error getting reference to the Map Layers Collection. Exiting. ")
                Exit Sub
            End If
        Catch ex As Exception
            MsgBox("Error getting reference to the Map Layers Collection. Exiting. " + ex.Message)
            Exit Sub
        End Try

        Dim pFeatureLayer As IFeatureLayer
        Dim pFeatureSelection As IFeatureSelection
        Dim pSelectionSet, pSelectionSetA As ISelectionSet
        Dim pFields As IFields
        Dim pWorkspaceA As IWorkspace
        Dim pTable As ITable
        Dim iFieldIndex As Integer
        ' FIPEX Custom Object: pworkspace, pfeaturelayer, fieldindex,pSelectionSet, FIPEX_EID  
        ' (helps to store all selection set details so that one edit session per unique worksapce can be used to speed things up)
        Dim pSelectAndUpdateFeaturesObject As SelectAndUpdateFeaturesObject = New SelectAndUpdateFeaturesObject(Nothing, Nothing, Nothing, _
                                                                                                                Nothing, Nothing, Nothing, Nothing)
        pEnumLayer.Reset()

        ' Look at the next layer in the list
        pFeatureLayer = CType(pEnumLayer.Next, IFeatureLayer)
        Do While Not pFeatureLayer Is Nothing ' these two lines must be separate
            If pFeatureLayer.Valid = True Then ' or there will be an empty object ref
                If pFeatureLayer.FeatureClass.ShapeType = esriGeometryType.esriGeometryLine Or _
                pFeatureLayer.FeatureClass.ShapeType = esriGeometryType.esriGeometryPolyline Or _
                pFeatureLayer.FeatureClass.ShapeType = esriGeometryType.esriGeometryMultipoint Or _
                pFeatureLayer.FeatureClass.ShapeType = esriGeometryType.esriGeometryPoint Then

                    pFeatureSelection = CType(pFeatureLayer, IFeatureSelection)
                    pSelectionSet = pFeatureSelection.SelectionSet

                    pFields = pFeatureLayer.FeatureClass.Fields

                    If pFeatureSelection.SelectionSet.Count <> 0 Then

                        pWorkspaceA = pFeatureLayer.FeatureClass.FeatureDataset.Workspace
                        pSelectionSetA = pFeatureSelection.SelectionSet

                        pTable = CType(pFeatureLayer.FeatureClass, ITable)
                        iFieldIndex = pTable.FindField(sLabelField)
                        If iFieldIndex = -1 Then
                            MsgBox("There must be a field named " & sLabelField & " in layer " & pFeatureLayer.Name & _
                                   ". Attempting to create one..")

                            ' GO Sep 2021 
                            CreateAttribute(sLabelField, pTable, pFeatureLayer)
                            pFields = pFeatureLayer.FeatureClass.Fields
                            pTable = CType(pFeatureLayer.FeatureClass, ITable)
                            iFieldIndex = pTable.FindField(sLabelField)
                            MsgBox("Successfully created field.")
                        End If

                        pSelectAndUpdateFeaturesObject = New SelectAndUpdateFeaturesObject(pWorkspaceA, pFeatureLayer, iFieldIndex, _
                                                                                           pSelectionSetA, Nothing, gEID, Nothing)
                        ' Can Predicate Search to get refined list of workspaces later
                        lSelectAndUpdateFeaturesObject.Add(pSelectAndUpdateFeaturesObject)
                    End If
                End If
            End If
            pFeatureLayer = CType(pEnumLayer.Next, IFeatureLayer)
        Loop
        pNetworkAnalysisExtResults.ClearResults()

    End Sub

    Public Shared Sub CreateAttribute(ByRef sLabelField As String, ByRef pTable As ITable, ByRef pFeatureLayer As IFeatureLayer)

        Dim pFieldEdit As IFieldEdit
        Dim pField As IField
        Dim pFields As IFields
        Dim pFieldsEdit As IFieldsEdit

        ' Create new Fields object
        pFields = New Fields
        pFieldsEdit = CType(pFields, IFieldsEdit)
        pFieldsEdit.FieldCount_2 = 1

        ' Create  Field
        pField = New Field
        pFieldEdit = CType(pField, IFieldEdit)
        pFieldEdit.AliasName_2 = sLabelField
        pFieldEdit.Name_2 = sLabelField
        pFieldEdit.Type_2 = esriFieldType.esriFieldTypeInteger

        pFieldsEdit.Field_2(0) = pField

        Try
            pFeatureLayer.FeatureClass.AddField(pFieldsEdit.Field(0))

        Catch ex As Exception
            MsgBox("Error adding field. " + ex.Message.ToString)
        End Try


    End Sub

    Public Shared Sub UpdateAttributesBatch(ByRef lSelectAndUpdateFeaturesObject As List(Of SelectAndUpdateFeaturesObject),
                                            ByRef sValueType As String)

        Dim lWorkspaces As List(Of ESRI.ArcGIS.Geodatabase.IWorkspace) = New List(Of IWorkspace)
        ' use findall to get the list of objects from master list containing each unique workspace
        Dim pWorkspace, pWorkspaceChecker As ESRI.ArcGIS.Geodatabase.IWorkspace
        Dim workspacecomparer As FindSelectionSetsByWorkspace2 'predicate
        Dim lRefinedSelectAndUpdateFeaturesObject As List(Of SelectAndUpdateFeaturesObject)
        Dim bWorkspaceMatch As Boolean = False
        Dim pFeatureLayer As IFeatureLayer

        Dim iFieldValue As Integer
        Dim dFieldValue As Double
        Dim sFieldValue As String
        Dim iFieldIndex As Integer
        Dim pSelectionSet2 As ESRI.ArcGIS.Geodatabase.ISelectionSet2
        Dim pFeatureCursor As IFeatureCursor
        Dim pFeature As IFeature
        Dim bError As Boolean = False
        Dim k As Integer
        Dim pEditor As ESRI.ArcGIS.Editor.IEditor
        Dim pID As New UID

        Try
            pID.Value = "{F8842F20-BB23-11D0-802B-0000F8037368}"
            pEditor = My.ArcMap.Application.FindExtensionByCLSID(pID)
            If pEditor Is Nothing Then
                MsgBox("Error getting reference to the Editor extension. Exiting. ")
                Exit Sub
            End If
        Catch ex As Exception
            MsgBox("Error getting reference to the Editor extension. Exiting. " + ex.Message)
            Exit Sub
        End Try

        For i = 0 To lSelectAndUpdateFeaturesObject.Count - 1

            ' check if workspace already edited using list tracker lWorkspaces
            bWorkspaceMatch = False
            pWorkspace = lSelectAndUpdateFeaturesObject(i).pWorkspace

            For j = 0 To lWorkspaces.Count - 1
                If lWorkspaces(j) Is pWorkspace Then
                    bWorkspaceMatch = True
                End If
            Next

            ' if it's not in the master list yet the updating hasn't been done yet
            If bWorkspaceMatch = False Then

                ' add it to the list
                lWorkspaces.Add(pWorkspace)

                ' get a refined master list for this workspace
                workspacecomparer = New FindSelectionSetsByWorkspace2(pWorkspace)
                lRefinedSelectAndUpdateFeaturesObject = New List(Of SelectAndUpdateFeaturesObject)
                lRefinedSelectAndUpdateFeaturesObject = lSelectAndUpdateFeaturesObject.FindAll(AddressOf workspacecomparer.CompareWorkspace)

                ' start one edit session for this refined list
                If pEditor.EditState = ESRI.ArcGIS.Editor.esriEditState.esriStateEditing Then
                    Try
                        pEditor.StopEditing(True)
                    Catch ex As Exception
                        MsgBox("Error trying to stop current edit session. Please save your edits and close this session and try again." + ex.Message)
                        Exit Sub
                    End Try
                End If
                Try
                    pEditor.StartEditing(pWorkspace)
                Catch ex As Exception
                    MsgBox("Error trying to start edit session for selected feature's feature class - layer " _
                           + ex.Message)
                End Try

                ' for all selection sets with the same workspace update their tables
                For j = 0 To lRefinedSelectAndUpdateFeaturesObject.Count - 1

                    ' just for safety check the workspaces
                    pWorkspaceChecker = lRefinedSelectAndUpdateFeaturesObject(j).pWorkspace
                    If Not pWorkspaceChecker Is pWorkspace Then
                        MsgBox("Error: the workspace in the refined master selection list isn't the one expected after predicate refining.")
                    End If

                    pEditor.StartOperation()
                    pFeatureCursor = Nothing

                    pSelectionSet2 = CType(lRefinedSelectAndUpdateFeaturesObject(j).pSelectionSet, ISelectionSet2)
                    pFeatureLayer = lRefinedSelectAndUpdateFeaturesObject(j).pFeatureLayer
                    iFieldIndex = lRefinedSelectAndUpdateFeaturesObject(j).iFieldIndex

                    ' flexibile types - GO Sep 2021
                    If sValueType = "integer" Then
                        iFieldValue = lRefinedSelectAndUpdateFeaturesObject(j).iValue
                    ElseIf sValueType = "string" Then
                        sFieldValue = lRefinedSelectAndUpdateFeaturesObject(j).sValue
                    ElseIf sValueType = "double" Then
                        dFieldValue = lRefinedSelectAndUpdateFeaturesObject(j).dValue
                    Else
                        iFieldValue = lRefinedSelectAndUpdateFeaturesObject(j).iValue
                    End If

                    Try
                        pSelectionSet2.Update(Nothing, True, pFeatureCursor)
                        pFeature = Nothing
                        pFeature = pFeatureCursor.NextFeature

                        Do Until pFeature Is Nothing
                            If sValueType = "integer" Then
                                pFeature.Value(iFieldIndex) = iFieldValue
                            ElseIf sValueType = "string" Then
                                pFeature.Value(iFieldIndex) = sFieldValue
                            ElseIf sValueType = "double" Then
                                pFeature.Value(iFieldIndex) = dFieldValue
                            Else
                                pFeature.Value(iFieldIndex) = iFieldValue
                            End If

                            pFeatureCursor.UpdateFeature(pFeature)
                            pFeature = pFeatureCursor.NextFeature
                        Loop

                        pEditor.StopOperation("featureupdated")
                    Catch ex As Exception
                        MsgBox("Error trying to update selected features. " + ex.Message _
                               + ". For layer " + pFeatureLayer.Name)
                    End Try
                Next ' selection object in the refined list

                ' end the edit sessions
                pEditor.StopEditing(True)

            End If
        Next ' i - object in the master list, looking for unique workspaces

    End Sub

    Public Shared Sub LabelFIPEXEIDs(ByRef pNodesList As IEnumNetEID,
                                      ByRef pFlagsList As IEnumNetEID,
                                      ByRef pUtilityNetworkAnalysisExt As ESRI.ArcGIS.EditorExt.IUtilityNetworkAnalysisExt,
                                      ByRef sFieldName As String,
                                      ByRef sValueType As String)

        ' GO Sep 2021
        ' No traces required, loop all barriers, branches, sinks, sources and label with EID.
        ' Label EIDs using their attribute tables

        ' for each node (barrier, branch, source)
        ' add to enum object

        ' for each flag 
        ' add to enum object

        ' create selection and store
        ' label

        Dim iEID, k As Integer
        Dim pEnumNetEIDGEN As IEnumNetEIDBuilderGEN = New EnumNetEIDArray
        Dim pEnumNetEID As IEnumNetEID
        Dim lSelectAndWriteObject As List(Of SelectAndUpdateFeaturesObject) = New List(Of SelectAndUpdateFeaturesObject)
        Dim pNetworkAnalysisExt As INetworkAnalysisExt
        Dim pNetworkAnalysisExtResults As INetworkAnalysisExtResults
        pNetworkAnalysisExt = CType(pUtilityNetworkAnalysisExt, INetworkAnalysisExt)
        pNetworkAnalysisExtResults = CType(pUtilityNetworkAnalysisExt, INetworkAnalysisExtResults)

        ' combine the flags and other nodes into one list
        k = 0
        pNodesList.Reset()

        For k = 0 To pNodesList.Count - 1
            iEID = pNodesList.Next()
            ' Need to send for selection one EID at a time, unfortunately
            pEnumNetEIDGEN = Nothing
            pEnumNetEIDGEN = New EnumNetEIDArray
            pEnumNetEIDGEN.Add(iEID)
            pEnumNetEID = Nothing
            pEnumNetEID = CType(pEnumNetEIDGEN, IEnumNetEID)
            ReturnSelectionSetList(iEID, sFieldName, lSelectAndWriteObject, pNetworkAnalysisExtResults, _
                               pNetworkAnalysisExt, pEnumNetEID, Nothing)
        Next
        k = 0
        pFlagsList.Reset()
        For k = 0 To pFlagsList.Count - 1
            iEID = pFlagsList.Next()
            ' Need to send for selection one EID at a time, unfortunately
            pEnumNetEIDGEN = Nothing
            pEnumNetEIDGEN = New EnumNetEIDArray
            pEnumNetEIDGEN.Add(iEID)
            pEnumNetEID = Nothing
            pEnumNetEID = CType(pEnumNetEIDGEN, IEnumNetEID)
            ReturnSelectionSetList(iEID, sFieldName, lSelectAndWriteObject, pNetworkAnalysisExtResults, _
                               pNetworkAnalysisExt, pEnumNetEID, Nothing)
        Next


        'MsgBox("The number of all nodes: " + Str(pNodesList.Count()))
        ' MsgBox("The number of all flags: " + Str(pFlagsList.Count()))
        'MsgBox("The number of all flags and other nodes combined: " + Str(pEnumNetEID.Count()))
        'MsgBox("The selection set of all nodes returned in LabelFIPEXEIDs: " + Str(lSelectAndWriteObject.Count))

        UpdateAttributesBatch(lSelectAndWriteObject, sValueType)

    End Sub


    Public Class FindSelectionSetsByWorkspace2
        ' predicate to retrieve matching workspace from a list of objects
        Private _pWorkspace As ESRI.ArcGIS.Geodatabase.IWorkspace

        Public Sub New(ByVal pworkspace As ESRI.ArcGIS.Geodatabase.IWorkspace)
            Me._pWorkspace = pworkspace
        End Sub

        Public Function CompareWorkspace(ByVal obj As SelectAndUpdateFeaturesObject) As Boolean
            Return (_pWorkspace Is obj.pWorkspace)
        End Function
    End Class

End Class

