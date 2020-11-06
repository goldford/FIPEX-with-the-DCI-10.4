
Imports ESRI.ArcGIS.ArcMapUI
Imports ESRI.ArcGIS.Carto
Imports ESRI.ArcGIS.Display
Imports ESRI.ArcGIS.EditorExt
Imports ESRI.ArcGIS.Geodatabase
Imports ESRI.ArcGIS.Geometry
Imports ESRI.ArcGIS.NetworkAnalysis
Imports ESRI.ArcGIS.esriSystem
Imports ESRI.ArcGIS.Editor

Public Class DistanceToMouth
    Inherits ESRI.ArcGIS.Desktop.AddIns.Button
    Private m_UNAExt As IUtilityNetworkAnalysisExt
    Private m_pNetworkAnalysisExt As INetworkAnalysisExt
    Private m_FiPEx__1 As FishPassageExtension
    Private m_sGeonetsettings, m_sGetNetworkLines As String
    Private m_iCurrentFeature, m_iTotalFeatures As Integer
    Friend DistanceForm As frmDistanceToMouth
    Friend ProgressForm As frmProgress_DistanceToMouth
    Dim pFindByFCIDPredicate As FindByFCIDPredicate
    Dim dTotalLength As Integer

    Public Sub New()

    End Sub

    Protected Overrides Sub OnClick()
        ' ===================================
        ' Distance To Mouth
        ' Created June 1, 2012
        ' Purpose: calculate distance to mouth / sink for 
        '          a large number of network lines / points
        '
        ' Created by: Greig Oldford
        '        For: Thesis
        ' 
        ' Description / Logic:
        '          Active Conditions - 
        '                             (1) network present
        '                             (2) some features selected
        '                             (3) network features editable (if adding field to them with results)
        '         Logic - 
        '                (1) pop-up form - describe tool to user
        '                (2) give option to use lines and/or points, add results to FC?, use FiPEx Options Quantity setting
        '                (3) pop-up progress form (updated each loop/flag)
        '                (4) get list of network features with selection present
        '                (5) save original network barriers and flags
        '                (6) clear all network flags and barriers
        '                (6) loop through point network features
        '                    - downstream total path trace (similar to 'advanced analysis' - ignores barriers)
        '

        ' =========================================
        ' get reference to the editor 

        ' need to add a field to this layer to 
        Dim uID As New UID
        uID.Value = "esriEditor.Editor"
        Dim pEditor As IEditor
        pEditor = CType(My.ArcMap.Application.FindExtensionByCLSID(uID), IEditor)
        If pEditor.EditState = esriEditState.esriStateEditing Then
            MsgBox("Please close all edit sessions before running Distance to Mouth/Source")
            Exit Sub
        End If

        ' =========================================
        ' Show User Form, get user settings

        DistanceForm = New frmDistanceToMouth
        Dim bCancel As Boolean = False
        Dim bUseFiPexQuan As Boolean
        Dim bSource, bMouth As Boolean
        ' list of flag layers, fcid, and field to update
        Dim lLayerNamesAndFCIDs As New List(Of LayersAndFCIDs)
        'Retrieve settings back from the form.
        Using DistanceForm As New frmDistanceToMouth 'm_sPOlyLayer was here!
            If DistanceForm.Form_Initialize(My.ArcMap.Application) Then
                DistanceForm.ShowDialog()
                bCancel = DistanceForm.m_bCancel1
                bUseFiPexQuan = DistanceForm.m_bUseFiPExQuan
                lLayerNamesAndFCIDs = DistanceForm.m_lLayersAndFCIDs
                bMouth = DistanceForm.chkMouth.Checked
                bSource = DistanceForm.chkSource.Checked
            End If
        End Using

        If bCancel Then
            Exit Sub
        End If

        If lLayerNamesAndFCIDs.Count = 0 Then
            MsgBox("You must select at least one network layer to use for start-points.")
            Exit Sub
        End If

        Dim i, j As Integer

        ' Test that user has permission to edit the feature classes by starting an edit session for their workspace
        i = 0
        For i = 0 To lLayerNamesAndFCIDs.Count - 1

        Next


        '' ===== Set up Progress Form ===========
        m_sGeonetsettings = "Save Current Geometric Network Settings"
        m_sGetNetworkLines = "Retrieve Current Network Edge Feature Classes and Length Fields"
        Dim poo As New Threading.Thread(AddressOf ManageResultsForm)
        'poo.Priority = System.Threading.ThreadPriority.Highest
        poo.Start()
        'Dim backgroundwerker As New System.ComponentModel.BackgroundWorker


        'End If
        'End Using


        ' =========================================
        ' get a list of network lines layers in the map
        Dim simpleEdgeFCs As IEnumFeatureClass
        Dim pGeometricNetwork As IGeometricNetwork = m_pNetworkAnalysisExt.CurrentNetwork
        simpleEdgeFCs = pGeometricNetwork.ClassesByType(ESRI.ArcGIS.Geodatabase.esriFeatureType.esriFTSimpleEdge)
        Dim pMap As IMap
        Dim pMxDoc As IMxDocument
        If pMxDoc Is Nothing Then
            pMxDoc = My.ArcMap.Application.Document
        End If
        pMap = pMxDoc.FocusMap
        Dim pFeatureLayer As IFeatureLayer
        Dim pFeatureClass As IFeatureClass


        Dim pLayerFCIDQuan As New LayerFCIDAndQuanField(Nothing, Nothing, Nothing)
        Dim lLayerFCIDQuan As New List(Of LayerFCIDAndQuanField)
        Dim pdataset As IDataset
        Dim slayername As String

        ' =========================================
        ' Get the quantity field for each simple network edge - 
        ' either from fipex settings or set as 'shape_length'
        m_sGetNetworkLines = "GETTING CURRENT NETWORK EDGE FEATURE CLASSES AND LENGTH FIELDS"
        If bUseFiPexQuan = True Then
            ' Get the FiPEx Quantities for the lines layer
            ' If settings have been set by the user then load them from the extension stream (stored in .mxd doc)
            If m_FiPEx__1.m_bLoaded = True Then

                ' get a list of included lines layers and their habitat length fields
                ' will need list of feature layer name, FCID, and length field names
                Dim iLinesCount As Integer
                iLinesCount = Convert.ToInt32(m_FiPEx__1.pPropset.GetProperty("numLines"))

                simpleEdgeFCs.Reset()
                pFeatureClass = simpleEdgeFCs.Next
                Do Until pFeatureClass Is Nothing

                    ' match any of the line layers saved in stream to those in listboxes
                    If iLinesCount > 0 Then
                        j = 0
                        For j = 0 To iLinesCount - 1
                            'sLineLayer = m_FiPEX__1.pPropset.GetProperty("IncLine" + j.ToString) ' get line layer
                            pLayerFCIDQuan = New LayerFCIDAndQuanField(Convert.ToString(m_FiPEx__1.pPropset.GetProperty("IncLine" + j.ToString)), Nothing, Nothing)
                            pdataset = DirectCast(pFeatureClass, IDataset)
                            slayername = pdataset.Name
                            If pLayerFCIDQuan._Layer = slayername Then
                                pLayerFCIDQuan._FCID = pFeatureClass.FeatureClassID
                                pLayerFCIDQuan._Quanfield = Convert.ToString(m_FiPEx__1.pPropset.GetProperty("LineQuanField" + j.ToString))
                                If pLayerFCIDQuan._Quanfield = "Not set" Then
                                    pLayerFCIDQuan._Quanfield = "Shape_Length"
                                End If
                                lLayerFCIDQuan.Add(pLayerFCIDQuan)
                            End If
                        Next
                    Else ' just assign the field as standard 'shape_length'
                        'sLineLayer = m_FiPEX__1.pPropset.GetProperty("IncLine" + j.ToString) ' get line layer
                        pLayerFCIDQuan = New LayerFCIDAndQuanField(pFeatureClass.LayerName, pFeatureClass.FeatureClassID, "Shape_Length")
                        lLayerFCIDQuan.Add(pLayerFCIDQuan)
                    End If
                    pFeatureClass = simpleEdgeFCs.Next
                Loop ' next simple edge

            Else
                bUseFiPexQuan = False
                simpleEdgeFCs.Reset()
                pFeatureClass = simpleEdgeFCs.Next
                Do Until pFeatureClass Is Nothing
                    pLayerFCIDQuan = New LayerFCIDAndQuanField(pFeatureClass.LayerName, pFeatureClass.FeatureClassID, "Shape_Length")
                    lLayerFCIDQuan.Add(pLayerFCIDQuan)
                    pFeatureClass = simpleEdgeFCs.Next
                Loop
            End If
        Else

            simpleEdgeFCs.Reset()
            pFeatureClass = simpleEdgeFCs.Next
            Do Until pFeatureClass Is Nothing
                pLayerFCIDQuan = New LayerFCIDAndQuanField(pFeatureClass.LayerName, pFeatureClass.FeatureClassID, "Shape_Length")
                lLayerFCIDQuan.Add(pLayerFCIDQuan)
                pFeatureClass = simpleEdgeFCs.Next
            Loop

        End If


        ' =========================================
        ' lLayerNamesAndFCIDs - list of layers using for source
        ' lLayerFCIDQuan - list of network lines and length fields

        m_sGetNetworkLines = "Retrieved Current Network Edge Feature Classes and Length Fields"
        m_sGeonetsettings = "SAVING CURRENT GEOMETRIC NETWORK SETTINGS"

        ' =========================================
        ' Save original network settings

        Dim pTraceTasks As ITraceTasks
        Dim pNetworkAnalysisExt As INetworkAnalysisExt
        Dim pNetworkAnalysisExtResults As INetworkAnalysisExtResults
        Dim pNetworkAnalysisExtBarriers As INetworkAnalysisExtBarriers
        Dim pOriginalJuncBarriersListGEN As IEnumNetEIDBuilderGEN
        Dim pOriginalEdgeBarriersListGEN As IEnumNetEIDBuilderGEN
        Dim pOriginalEdgeFlagsListGEN As IEnumNetEIDBuilderGEN
        Dim pOriginaljuncFlagsListGEN As IEnumNetEIDBuilderGEN
        Dim pOriginalJuncBarriersList As IEnumNetEID
        Dim pOriginalEdgeBarriersList As IEnumNetEID
        Dim pOriginalEdgeFlagsList As IEnumNetEID
        Dim pOriginalJuncFlagsList As IEnumNetEID
        Dim pNetwork As INetwork
        Dim pNetElements As INetElements
        Dim pNetworkAnalysisExtFlags As INetworkAnalysisExtFlags
        Dim bFlagDisplay As IFlagDisplay
        Dim bEID As Integer

        ' =============== SAVE ORIGINAL GEONet SETTINGS =========================
        pNetworkAnalysisExt = CType(m_UNAExt, INetworkAnalysisExt)
        pGeometricNetwork = pNetworkAnalysisExt.CurrentNetwork

        pNetwork = pGeometricNetwork.Network
        pNetElements = CType(pNetwork, INetElements)

        ' QI for the ITraceTasks interface using IUtilityAnalysisExt
        pTraceTasks = CType(m_UNAExt, ITraceTasks)

        ' update extension with the results
        ' QI for the INetworkAnalysisExtResults interface using IUTilityNetworkAnalysisExt
        pNetworkAnalysisExtResults = CType(m_UNAExt, INetworkAnalysisExtResults)

        ' clear any leftover results from previous calls to the cmd
        pNetworkAnalysisExtResults.ClearResults()

        ' QI the Flags and barriers
        pNetworkAnalysisExtFlags = CType(m_UNAExt, INetworkAnalysisExtFlags)
        pNetworkAnalysisExtBarriers = CType(m_UNAExt, INetworkAnalysisExtBarriers)

        pOriginalEdgeBarriersListGEN = New EnumNetEIDArray
        pOriginalJuncBarriersListGEN = New EnumNetEIDArray
        pOriginalEdgeFlagsListGEN = New EnumNetEIDArray
        pOriginaljuncFlagsListGEN = New EnumNetEIDArray

        Dim pEdgeFlagDisplayTEMP As IEdgeFlagDisplay
        Dim lEdgeBarrierDisplaylist As New List(Of IEdgeFlagDisplay)
        Dim lEdgeFlagDisplaylist As New List(Of IEdgeFlagDisplay)

        ' -----------------------------------------
        ' Save the junction barriers
        For i = 0 To pNetworkAnalysisExtBarriers.JunctionBarrierCount - 1
            ' Use bFlagDisplay to retrieve EIDs of the barriers for later
            bFlagDisplay = CType(pNetworkAnalysisExtBarriers.JunctionBarrier(i), IFlagDisplay)
            bEID = pNetElements.GetEID(bFlagDisplay.FeatureClassID, bFlagDisplay.FID, bFlagDisplay.SubID, esriElementType.esriETJunction)
            pOriginalJuncBarriersListGEN.Add(bEID)
        Next

        ' QI to and get an array object that has 'count' and 'next' methods
        pOriginalJuncBarriersList = CType(pOriginalJuncBarriersListGEN, IEnumNetEID)

        ' -----------------------------------------
        ' Save the edge barriers
        i = 0
        For i = 0 To pNetworkAnalysisExtBarriers.EdgeBarrierCount - 1
            ' Use bFlagDisplay to retrieve EIDs of the barriers for later
            pEdgeFlagDisplayTEMP = pNetworkAnalysisExtBarriers.EdgeBarrier(i)

            ' should get rid of below code because it doesn't store geometry of edge flag (x,y position along line)
            bEID = pNetElements.GetEID(bFlagDisplay.FeatureClassID, bFlagDisplay.FID, bFlagDisplay.SubID, esriElementType.esriETEdge)
            pOriginalEdgeBarriersListGEN.Add(bEID)

            ' Add the flagdisplay to the list
            lEdgeBarrierDisplaylist.Add(pEdgeFlagDisplayTEMP)

        Next

        ' QI to and get an array object that has 'count' and 'next' methods
        pOriginalEdgeBarriersList = CType(pOriginalEdgeBarriersListGEN, IEnumNetEID)

        ' -----------------------------------------
        ' Save the junction flags
        i = 0
        For i = 0 To pNetworkAnalysisExtFlags.JunctionFlagCount - 1
            ' Use the bFlagDisplay to retrieve the EIDs of the junction flags
            bFlagDisplay = CType(pNetworkAnalysisExtFlags.JunctionFlag(i), IFlagDisplay)
            bEID = pNetElements.GetEID(bFlagDisplay.FeatureClassID, bFlagDisplay.FID, bFlagDisplay.SubID, esriElementType.esriETJunction)
            pOriginaljuncFlagsListGEN.Add(bEID)
        Next

        ' QI to and get an array interface that has 'count' and 'next' methods
        'pOriginaljuncFlagsList = 
        pOriginalJuncFlagsList = CType(pOriginaljuncFlagsListGEN, IEnumNetEID)

        ' -----------------------------------------
        ' save the edge flags
        i = 0
        For i = 0 To pNetworkAnalysisExtFlags.EdgeFlagCount - 1

            ' Use the bFlagDisplay to retrieve EIDs of the Edge flags for later
            pEdgeFlagDisplayTEMP = pNetworkAnalysisExtFlags.EdgeFlag(i)

            ' should get rid of the below 2 lines of  code
            bEID = pNetElements.GetEID(bFlagDisplay.FeatureClassID, bFlagDisplay.FID, bFlagDisplay.SubID, esriElementType.esriETEdge)
            pOriginalEdgeFlagsListGEN.Add(bEID)

            ' add the edgedisplay to the list
            lEdgeFlagDisplaylist.Add(pEdgeFlagDisplayTEMP)

        Next

        ' QI to and get an array interface that has 'count' and 'next' methods
        pOriginalEdgeFlagsList = CType(pOriginalEdgeFlagsListGEN, IEnumNetEID)
        ' ******************************************

        pNetworkAnalysisExtFlags.ClearFlags()
        pNetworkAnalysisExtBarriers.ClearBarriers()

        Dim eFlowElements As esriFlowElements
        Dim pEnumNetEIDBuilder As IEnumNetEIDBuilder
        Dim pTraceFlowSolver As ITraceFlowSolver
        Dim pResultEdges As IEnumNetEID
        Dim pResultJunctions As IEnumNetEID
        Dim barrierCount As Integer

        ' look for shape_length
        ' Dim dblLength as double
        'Dim pCurve As ICurve

        ' ICurve curve = feat.Shape as ICurve;
        'pCurve = [shape]
        'dblLength = pCurve.Length

        m_sGeonetsettings = "Saved Current Geometric Network Settings"
        'ProgressForm.Close()



        ' lLayerNamesAndFCIDs - list of layers using for source
        ' lLayerFCIDQuan - list of network lines and length fields

        ' =====================================================
        ' Loop through each of the layers using for source. 
        ' If junction layer, place junction flags, 
        ' if edge layer place edge flags
        Dim bEdge As Boolean = False
        Dim iIndex As Integer = 0
        Dim pLayerFCIDAndHabQuan As New LayerFCIDAndHabQuan(Nothing, Nothing, Nothing)
        Dim lLayerFCIDAndHabQuan As New List(Of LayerFCIDAndHabQuan)
        Dim dRunningQuanTally As Double = 0.0
        Try
            Dim pFeatureSelection As IFeatureSelection
            Dim pSelectionSet As ISelectionSet
            Dim pFields As IFields
            Dim iClassCheckTemp As Integer
            Dim pEnumIDs As IEnumIDs
            Dim selectedFID, selectedFCID As Integer
            Dim pEID As IEnumNetEID
            Dim iEID, iEIDCount, iFCID, iFID, iSubID As Integer
            Dim pSymbol As ISymbol
            Dim pFlagDisplay As IFlagDisplay
            Dim pJunctionFlagDisplay As IJunctionFlagDisplay
            Dim pEdgeFlagDisplay As IEdgeFlagDisplay
            Dim pActiveView As IActiveView = CType(pMap, IActiveView)

            Dim pGeometry As ESRI.ArcGIS.Geometry.IGeometry
            Dim pPolyline As ESRI.ArcGIS.Geometry.IPolyline
            Dim pFeature As IFeature
            Dim pFeatureType As ESRI.ArcGIS.Geodatabase.esriFeatureType
            Dim pPoint As ESRI.ArcGIS.Geometry.IPoint = New ESRI.ArcGIS.Geometry.Point


            ' =================================================
            ' set up the symbol used for flags
            ' Flag symbol code
            Dim pFlagSymbol As ESRI.ArcGIS.Display.ISimpleMarkerSymbol
            ' Create simple marker symbol for barriers
            Dim pSimpleMarkerSymbol As ESRI.ArcGIS.Display.ISimpleMarkerSymbol
            Dim pRgbColor As ESRI.ArcGIS.Display.IRgbColor

            pRgbColor = New ESRI.ArcGIS.Display.RgbColor
            With pRgbColor
                .Red = 0
                .Green = 255
                .Blue = 0
            End With
            pSimpleMarkerSymbol = New ESRI.ArcGIS.Display.SimpleMarkerSymbol
            With pSimpleMarkerSymbol
                .Color = pRgbColor
                .Style = ESRI.ArcGIS.Display.esriSimpleMarkerStyle.esriSMSSquare
                .Outline = True
                .Size = 10
            End With

            ' Result is a global variable containing a barrier marker
            pFlagSymbol = pSimpleMarkerSymbol

            Dim pFLyrSlct As IFeatureLayer
            Dim pUID As New UID
            Dim pEnumLayer As IEnumLayer
            Dim iResultEID, iResultFCID, iResultFID, iResultSubID As Integer
            Dim pResultsIDsObject As New ResultsIDsObject(Nothing, Nothing, Nothing, Nothing) ' custom object for results ids
            Dim lResultsIDsObject As New List(Of ResultsIDsObject)
            Dim lRefinedResultsIDsObject As New List(Of ResultsIDsObject)
            Dim pResultsIDObjectComparer As FindByFCIDPredicate
            Dim iMapFCID, iMapFID, iMapSubID As Integer ' map TOC layer id's
            Dim pFindByFCIDPredicate2 As FindByFCIDPredicate2
            Dim iIndex2, k As Integer
            Dim bMatch As Boolean
            Dim bFieldFound As Boolean
            Dim iFieldIndex As Integer
            Dim vTemp As Object
            Dim dTemp As Double
            Dim sFieldNameMouth As String = "DistanceToMouth"
            Dim sFieldNameMouth2 As String
            Dim sFieldNameSource As String = "DistanceToSource"
            Dim sFieldNameSource2 As String
            Dim pField As IField ' Create ObjectID Field
            Dim pFieldEdit As IFieldEdit

            Dim iFieldMatch As Integer = 0
            Dim lFCID_FID_dTotal_Mouth As New List(Of FCID_FID_dTotalObject)
            Dim pFCID_FID_dTotal_Mouth As New FCID_FID_dTotalObject(Nothing, Nothing, Nothing)
            Dim lFCID_FID_dTotal_Source As New List(Of FCID_FID_dTotalObject)
            Dim pFCID_FID_dTotal_Source As New FCID_FID_dTotalObject(Nothing, Nothing, Nothing)
            Dim pFCID_FID_dTotalObject_Predicate As FCID_FID_dTotalObject_Predicate
            Dim pQueryFilter As IQueryFilter
            Dim pFeatureCursor As IFeatureCursor
            Dim sOIDName As String
            '=========================================
            ' Loop through each source element layer, place flag, perform trace

            i = 0
            For i = 0 To lLayerNamesAndFCIDs.Count - 1


                ' Start a new list of distance to mouth for this layer
                lFCID_FID_dTotal_Mouth = New List(Of FCID_FID_dTotalObject)
                lFCID_FID_dTotal_Source = New List(Of FCID_FID_dTotalObject)


                ' ------------------------------------------------
                ' loop through map till find the map index of layer

                j = 0
                For j = 0 To pMap.LayerCount - 1
                    If pMap.Layer(j).Valid = True Then
                        ' If it's a feature layer then
                        If TypeOf pMap.Layer(j) Is IFeatureLayer Then
                            pFeatureLayer = CType(pMap.Layer(j), IFeatureLayer)
                            ' If it's a junction then set the type as a simple edge
                            If pFeatureLayer.Name = lLayerNamesAndFCIDs(i).LayerName Then
                                iIndex = j
                                If pFeatureLayer.FeatureClass.FeatureType = esriFeatureType.esriFTSimpleEdge Then
                                    bEdge = True
                                ElseIf pFeatureLayer.FeatureClass.FeatureType = esriFeatureType.esriFTSimpleJunction Then
                                    bEdge = False
                                End If
                                Exit For
                            End If
                        End If
                    End If
                Next

                ' got edge/junction and index of layer
                ' set a edge / junction flag on each selected feature and do the trace.  
                ' update object with layer, layerID, and total path downstream quantity


                ' ======== ADD OUTPUT FIELDS ==============

                If bMouth = True Then

                    ' =========================================
                    ' Make sure field name does NOT already exist

                    pFeatureLayer = CType(pMap.Layer(iIndex), IFeatureLayer)
                    iFieldMatch = pFeatureLayer.FeatureClass.FindField(sFieldNameMouth)
                    If iFieldMatch <> -1 Then
                        j = 0
                        Do Until iFieldMatch = -1
                            j += 1
                            sFieldNameMouth2 = sFieldNameMouth + Convert.ToString(j)
                            iFieldMatch = pFeatureLayer.FeatureClass.FindField(sFieldNameMouth2)
                        Loop
                        sFieldNameMouth = sFieldNameMouth2
                    End If

                    ' Add the field to hold the output to this layer
                    Try

                        pField = New Field
                        pFieldEdit = CType(pField, IFieldEdit)
                        pFieldEdit.AliasName_2 = sFieldNameMouth
                        pFieldEdit.Name_2 = sFieldNameMouth
                        pFieldEdit.Editable_2 = True
                        pFieldEdit.Type_2 = esriFieldType.esriFieldTypeDouble
                        pField = CType(pFieldEdit, IField)
                        pFeatureLayer.FeatureClass.AddField(pField)

                    Catch ex As Exception
                        MsgBox("Failed adding new field to feautre class. " & ex.Message)
                    End Try

                End If 'bmouth is true

                If bSource = True Then

                    ' =========================================
                    ' Make sure field name does NOT already exist

                    pFeatureLayer = CType(pMap.Layer(iIndex), IFeatureLayer)
                    iFieldMatch = pFeatureLayer.FeatureClass.FindField(sFieldNameSource)
                    If iFieldMatch <> -1 Then
                        j = 0
                        Do Until iFieldMatch = -1
                            j += 1
                            sFieldNameSource2 = sFieldNameSource + Convert.ToString(j)
                            iFieldMatch = pFeatureLayer.FeatureClass.FindField(sFieldNameSource2)
                        Loop
                        sFieldNameSource = sFieldNameSource2
                    End If

                    ' Add the field to hold the output to this layer
                    Try

                        pField = New Field
                        pFieldEdit = CType(pField, IFieldEdit)
                        pFieldEdit.AliasName_2 = sFieldNameSource
                        pFieldEdit.Name_2 = sFieldNameSource
                        pFieldEdit.Editable_2 = True
                        pFieldEdit.Type_2 = esriFieldType.esriFieldTypeDouble
                        pField = CType(pFieldEdit, IField)
                        pFeatureLayer.FeatureClass.AddField(pField)

                    Catch ex As Exception
                        MsgBox("Failed adding new field to feature class. " & ex.Message)
                    End Try

                End If 'bsource is true

                ' ===============================================
                ' Get a list of EIDs for each of the selected features
                pFeatureSelection = CType(pFeatureLayer, IFeatureSelection)
                pEnumIDs = pFeatureSelection.SelectionSet.IDs
                pEnumIDs.Reset()
                selectedFID = pEnumIDs.Next
                selectedFCID = pFeatureLayer.FeatureClass.FeatureClassID
                pNetElements = CType(pNetwork, INetElements)
                pNetworkAnalysisExtBarriers = CType(m_UNAExt, INetworkAnalysisExtBarriers)

                ' update total features selected for this fc
                m_iTotalFeatures = pFeatureSelection.SelectionSet.Count
                m_iCurrentFeature = 0
                ' Loop through each selected feature and place a flag and run analysis
                Do Until selectedFID < 0

                    m_iCurrentFeature = m_iCurrentFeature + 1

                    ' ------------------------------------
                    ' Set a flag on the selected feature

                    ' Get corresponding EID(s) of Feature Selection ID
                    If bEdge = True Then
                        ' for some reason without a feature 'subid' can't use the alternative GETEID method. 
                        ' so this returns an array - the array should always be of length 1
                        pEID = pNetElements.GetEIDs(selectedFCID, selectedFID, ESRI.ArcGIS.Geodatabase.esriElementType.esriETEdge)
                    Else
                        pEID = pNetElements.GetEIDs(selectedFCID, selectedFID, ESRI.ArcGIS.Geodatabase.esriElementType.esriETJunction)
                    End If

                    If Not pEID Is Nothing Then
                        pEID.Reset()
                        ' Get a count of element IDs
                        iEIDCount = pEID.Count

                        If iEIDCount > 1 Then
                            MsgBox("Warning: Call to Element ID of source layer " + pFeatureLayer.Name + " returned more than one net element. Trace results may be inaccurate.")
                        End If

                        iEID = pEID.Next
                    Else
                        MsgBox("Warning: No Element ID Returned for first selected feature in " + pFeatureLayer.Name + ". The trace results may be incomplete.")
                        Exit Do
                    End If

                    ' ===========================================================================
                    ' Set the flag

                    ' Query the corresponding user ID's to the element ID
                    If bEdge = True Then
                        pNetElements.QueryIDs(iEID, ESRI.ArcGIS.Geodatabase.esriElementType.esriETEdge, iFCID, iFID, iSubID)
                    Else
                        pNetElements.QueryIDs(iEID, ESRI.ArcGIS.Geodatabase.esriElementType.esriETJunction, iFCID, iFID, iSubID)
                    End If

                    'lFIDarray(p) = lFID
                    'pSelectionSetBarriers.Add(lFCID, lFID)

                    pSymbol = CType(pFlagSymbol, ISymbol)

                    ' Display the barriers as a Junction or Edge FlagDisplay type
                    If bEdge = True Then
                        pFlagDisplay = New EdgeFlagDisplay

                        'Dim pEdgeFlagDisplay As IEdgeFlagDisplay
                        'pEdgeFlagDisplay.Percentage = 0.5

                        'pGeometry = pGeometricNetwork.GeometryForEdgeEID(iEID)
                        'Dim pEdgeFlag As EdgeFlag
                        'Dim pFlagDisplay As IFlagDisplay
                        pFeature = pFeatureLayer.FeatureClass.GetFeature(selectedFID)
                        pPolyline = pFeature.Shape
                        pPolyline.QueryPoint(ESRI.ArcGIS.Geometry.esriSegmentExtension.esriExtendAtTo, 0.5, True, pPoint)
                        pGeometry = CType(pPoint, ESRI.ArcGIS.Geometry.IGeometry)

                        With pFlagDisplay
                            .FeatureClassID = iFCID
                            .FID = iFID
                            .Geometry = pGeometry
                            .Symbol = pSymbol
                        End With
                    Else
                        pFlagDisplay = New JunctionFlagDisplay
                        With pFlagDisplay
                            .FeatureClassID = iFCID
                            .FID = iFID
                            .Geometry = pGeometricNetwork.GeometryForJunctionEID(iEID)
                            .Symbol = pSymbol
                        End With
                    End If

                    ' Add the barriers to the logical network
                    If bEdge = True Then
                        pEdgeFlagDisplay = CType(pFlagDisplay, IEdgeFlagDisplay)
                        pNetworkAnalysisExtFlags.AddEdgeFlag(pEdgeFlagDisplay)
                    Else
                        pJunctionFlagDisplay = CType(pFlagDisplay, IJunctionFlagDisplay)
                        pNetworkAnalysisExtFlags.AddJunctionFlag(pJunctionFlagDisplay)
                    End If

                    pActiveView.Refresh()

                    '' ------------------------------------
                    '' set all layers as selectable 
                    '' Make sure all the layers in the TOC are selectable
                    'j = 0
                    'For j = 0 To pMap.LayerCount - 1
                    '    If pMap.Layer(j).Valid = True Then
                    '        If TypeOf pMap.Layer(j) Is IFeatureLayer Then
                    '            pFLyrSlct = CType(pMap.Layer(j), IFeatureLayer)
                    '            pFLyrSlct.Selectable = True
                    '        End If
                    '    End If
                    'Next


                    ' ------------------------------------
                    ' perform trace

                    ' ====================== RUN TRACE IN DIRECTION OF ANALYSIS ====================
                    '                            TO GET FLOW END ELEMENTS
                    If bMouth = True Then

                        pNetworkAnalysisExtResults.ClearResults()

                        'prepare the network solver
                        pTraceFlowSolver = TraceFlowSolverSetup3()
                        If pTraceFlowSolver Is Nothing Then
                            System.Windows.Forms.MessageBox.Show("Could not set up the network. Check that there is a network loaded.", "TraceFlowSolver setup error.")
                            Exit Sub
                        End If

                        eFlowElements = esriFlowElements.esriFEJunctionsAndEdges
                        pTraceFlowSolver.FindFlowElements(esriFlowMethod.esriFMDownstream, eFlowElements, pResultJunctions, pResultEdges)

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

                        ' save all the eid elements returned to a list, but retrieve their FCIDs, FIDs, etc. first
                        ' pResultsIDsObject, lResultsIDsObject, lRefinedResultsIDsObject, FindByFCIDPredicate
                        pResultEdges.Reset()
                        j = 0
                        lResultsIDsObject = New List(Of ResultsIDsObject)
                        For j = 0 To pResultEdges.Count - 1

                            iResultEID = pResultEdges.Next
                            pNetElements.QueryIDs(iResultEID, esriElementType.esriETEdge, iResultFCID, iResultFID, iResultSubID)
                            pResultsIDsObject = New ResultsIDsObject(iResultEID, iResultFCID, iResultFID, iResultSubID)
                            lResultsIDsObject.Add(pResultsIDsObject)
                        Next ' j


                        '' ============ SAVE line or point feature FOR HIGHLIGHTING ==================
                        '' Get results to display as highlights at end of sub
                        'pResultJunctions.Reset()
                        'k = 0
                        'For k = 0 To pResultJunctions.Count - 1
                        '    pTotalResultsJunctionsGEN.Add(pResultJunctions.Next)
                        'Next
                        'pResultEdges.Reset()
                        'For k = 0 To pResultEdges.Count - 1
                        '    pTotalResultsEdgesGEN.Add(pResultEdges.Next)
                        'Next


                        '' Get results as selection
                        '' Can't do this because original selection of features to trace from would be erased
                        'pMap.ClearSelection()
                        'pNetworkAnalysisExtResults.CreateSelection(pResultJunctions, pResultEdges)


                        '' ---- EXCLUDE FEATURES -----
                        'pEnumLayer.Reset()
                        '' Look at the next layer in the list
                        'pFeatureLayer = CType(pEnumLayer.Next, IFeatureLayer)
                        'Do While Not pFeatureLayer Is Nothing ' these two lines must be separate
                        '    If pFeatureLayer.Valid = True Then
                        '        If pFeatureLayer.FeatureClass.ShapeType = esriGeometryType.esriGeometryLine Or _
                        '        pFeatureLayer.FeatureClass.ShapeType = esriGeometryType.esriGeometryPolygon Or _
                        '        pFeatureLayer.FeatureClass.ShapeType = esriGeometryType.esriGeometryPolyline Then ' or there will be an empty object ref
                        '            ExcludeFeatures(pFeatureLayer)
                        '        End If
                        '    End If
                        '    pFeatureLayer = CType(pEnumLayer.Next, IFeatureLayer)
                        'Loop
                        '' ---- END EXCLUDE FEATURES -----


                        ' ------------------------------------
                        ' calculate stats from selection return

                        ' loop through selected features 
                        ' and add up the habitat quantity field

                        ' lLayerNamesAndFCIDs - list of layers using for source
                        ' lLayerFCIDQuan - list of network lines and length fields
                        '
                        ' FindByFCIDPredicate - to refine list of result IDs object

                        pUID = New UID
                        pUID.Value = "{E156D7E5-22AF-11D3-9F99-00C04F6BC78E}"

                        pEnumLayer = pMap.Layers(pUID, True)


                        ' For each of the lines layers included in this analysis
                        j = 0

                        dTotalLength = 0
                        For j = 0 To lLayerFCIDQuan.Count - 1


                            ' for each map layer in the map
                            pEnumLayer.Reset()
                            pFeatureLayer = CType(pEnumLayer.Next, IFeatureLayer)


                            ' the match var is used to exit the loop of the map layers as soon 
                            ' as a match is encountered.  this is in case the network layer is present 
                            ' in the map twice (prevent double counting)
                            bMatch = False
                            Do While Not pFeatureLayer Is Nothing ' these two lines must be separate
                                If pFeatureLayer.Valid = True Then ' or there will be an empty object ref

                                    iMapFCID = pFeatureLayer.FeatureClass.FeatureClassID

                                    If iMapFCID = lLayerFCIDQuan(j)._FCID Then
                                        bMatch = True
                                        Exit Do
                                    End If

                                    pFeatureLayer = CType(pEnumLayer.Next, IFeatureLayer)

                                End If 'valid' feature
                            Loop

                            ' If there was a match found in the TOC then
                            ' find that layer and get a refined list of results for that FCID
                            ' and then loop through those results and add them to a total amount. 

                            If bMatch = True Then

                                ' get a refined list from the results
                                ' pResultsIDsObject, lResultsIDsObject, lRefinedResultsIDsObject, FindByFCIDPredicate

                                pResultsIDObjectComparer = New FindByFCIDPredicate(iMapFCID)
                                lRefinedResultsIDsObject = lResultsIDsObject.FindAll(AddressOf pResultsIDObjectComparer.CompareByFCID)

                                ' if there are results for this feature in the map TOC
                                If lRefinedResultsIDsObject.Count > 0 Then

                                    Try

                                        ' get the fields for this feature class
                                        pFields = pFeatureLayer.FeatureClass.Fields

                                        ' Get the index of the length quantity field from the list of fields 
                                        ' for this featureclass
                                        iFieldIndex = pFields.FindField(lLayerFCIDQuan(j)._Quanfield)
                                        If iFieldIndex <> -1 Then
                                            bFieldFound = True

                                            ' Get length quantity from each of the results features in the list for 
                                            ' this map layer.
                                            k = 0
                                            For k = 0 To lRefinedResultsIDsObject.Count - 1

                                                pFeature = pFeatureLayer.FeatureClass.GetFeature(lRefinedResultsIDsObject(k)._FID)
                                                vTemp = pFeature.Value(iFieldIndex)
                                                Try
                                                    dTemp = Convert.ToDouble(vTemp)
                                                Catch ex As Exception
                                                    MsgBox("Can't convert value in DistanceToMouth calculation retrieved from field to type 'double' (decimal) for feature class: " & pFeatureLayer.FeatureClass.AliasName)
                                                End Try

                                                dTotalLength = dTemp + dTotalLength
                                            Next


                                            ' For this feature / flag have found the total length now, for this feature class.  
                                            ' if there are other feature classes continue to loop through

                                        Else
                                            bFieldFound = False
                                            MsgBox("Can't find the field for this feature class, even though it's a network line... feauture class: " & pFeatureLayer.FeatureClass.AliasName)
                                        End If

                                    Catch ex As Exception
                                        MsgBox("Error encountered in DistanceToMouth during retrieval of fields or values from feature class: " & pFeatureLayer.FeatureClass.AliasName & ex.Message)

                                    End Try
                                End If
                            End If

                        Next ' edge network feature class

                        pFCID_FID_dTotal_Mouth = New FCID_FID_dTotalObject(selectedFCID, selectedFID, dTotalLength)
                        lFCID_FID_dTotal_Mouth.Add(pFCID_FID_dTotal_Mouth)
                        dTotalLength = 0

                    End If 'bMouth is true

                    If bSource = True Then
                        pNetworkAnalysisExtResults.ClearResults()

                        'prepare the network solver
                        pTraceFlowSolver = TraceFlowSolverSetup3()
                        If pTraceFlowSolver Is Nothing Then
                            System.Windows.Forms.MessageBox.Show("Could not set up the network. Check that there is a network loaded.", "TraceFlowSolver setup error.")
                            Exit Sub
                        End If

                        eFlowElements = esriFlowElements.esriFEJunctionsAndEdges
                        pTraceFlowSolver.FindFlowElements(esriFlowMethod.esriFMUpstream, eFlowElements, pResultJunctions, pResultEdges)

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

                        ' save all the eid elements returned to a list, but retrieve their FCIDs, FIDs, etc. first
                        ' pResultsIDsObject, lResultsIDsObject, lRefinedResultsIDsObject, FindByFCIDPredicate
                        pResultEdges.Reset()
                        j = 0
                        lResultsIDsObject = New List(Of ResultsIDsObject)
                        For j = 0 To pResultEdges.Count - 1

                            iResultEID = pResultEdges.Next
                            pNetElements.QueryIDs(iResultEID, esriElementType.esriETEdge, iResultFCID, iResultFID, iResultSubID)
                            pResultsIDsObject = New ResultsIDsObject(iResultEID, iResultFCID, iResultFID, iResultSubID)
                            lResultsIDsObject.Add(pResultsIDsObject)
                        Next ' j


                        '' ============ SAVE line or point feature FOR HIGHLIGHTING ==================
                        '' Get results to display as highlights at end of sub
                        'pResultJunctions.Reset()
                        'k = 0
                        'For k = 0 To pResultJunctions.Count - 1
                        '    pTotalResultsJunctionsGEN.Add(pResultJunctions.Next)
                        'Next
                        'pResultEdges.Reset()
                        'For k = 0 To pResultEdges.Count - 1
                        '    pTotalResultsEdgesGEN.Add(pResultEdges.Next)
                        'Next


                        '' Get results as selection
                        '' Can't do this because original selection of features to trace from would be erased
                        'pMap.ClearSelection()
                        'pNetworkAnalysisExtResults.CreateSelection(pResultJunctions, pResultEdges)


                        '' ---- EXCLUDE FEATURES -----
                        'pEnumLayer.Reset()
                        '' Look at the next layer in the list
                        'pFeatureLayer = CType(pEnumLayer.Next, IFeatureLayer)
                        'Do While Not pFeatureLayer Is Nothing ' these two lines must be separate
                        '    If pFeatureLayer.Valid = True Then
                        '        If pFeatureLayer.FeatureClass.ShapeType = esriGeometryType.esriGeometryLine Or _
                        '        pFeatureLayer.FeatureClass.ShapeType = esriGeometryType.esriGeometryPolygon Or _
                        '        pFeatureLayer.FeatureClass.ShapeType = esriGeometryType.esriGeometryPolyline Then ' or there will be an empty object ref
                        '            ExcludeFeatures(pFeatureLayer)
                        '        End If
                        '    End If
                        '    pFeatureLayer = CType(pEnumLayer.Next, IFeatureLayer)
                        'Loop
                        '' ---- END EXCLUDE FEATURES -----


                        ' ------------------------------------
                        ' calculate stats from selection return

                        ' loop through selected features 
                        ' and add up the habitat quantity field

                        ' lLayerNamesAndFCIDs - list of layers using for source
                        ' lLayerFCIDQuan - list of network lines and length fields
                        '
                        ' FindByFCIDPredicate - to refine list of result IDs object

                        pUID = New UID
                        pUID.Value = "{E156D7E5-22AF-11D3-9F99-00C04F6BC78E}"

                        pEnumLayer = pMap.Layers(pUID, True)


                        ' For each of the lines layers included in this analysis
                        j = 0

                        dTotalLength = 0
                        For j = 0 To lLayerFCIDQuan.Count - 1


                            ' for each map layer in the map
                            pEnumLayer.Reset()
                            pFeatureLayer = CType(pEnumLayer.Next, IFeatureLayer)


                            ' the match var is used to exit the loop of the map layers as soon 
                            ' as a match is encountered.  this is in case the network layer is present 
                            ' in the map twice (prevent double counting)
                            bMatch = False
                            Do While Not pFeatureLayer Is Nothing ' these two lines must be separate
                                If pFeatureLayer.Valid = True Then ' or there will be an empty object ref

                                    iMapFCID = pFeatureLayer.FeatureClass.FeatureClassID

                                    If iMapFCID = lLayerFCIDQuan(j)._FCID Then
                                        bMatch = True
                                        Exit Do
                                    End If

                                    pFeatureLayer = CType(pEnumLayer.Next, IFeatureLayer)

                                End If 'valid' feature
                            Loop

                            ' If there was a match found in the TOC then
                            ' find that layer and get a refined list of results for that FCID
                            ' and then loop through those results and add them to a total amount. 

                            If bMatch = True Then

                                ' get a refined list from the results
                                ' pResultsIDsObject, lResultsIDsObject, lRefinedResultsIDsObject, FindByFCIDPredicate

                                pResultsIDObjectComparer = New FindByFCIDPredicate(iMapFCID)
                                lRefinedResultsIDsObject = lResultsIDsObject.FindAll(AddressOf pResultsIDObjectComparer.CompareByFCID)

                                ' if there are results for this feature in the map TOC
                                If lRefinedResultsIDsObject.Count > 0 Then

                                    Try

                                        ' get the fields for this feature class
                                        pFields = pFeatureLayer.FeatureClass.Fields

                                        ' Get the index of the length quantity field from the list of fields 
                                        ' for this featureclass
                                        iFieldIndex = pFields.FindField(lLayerFCIDQuan(j)._Quanfield)
                                        If iFieldIndex <> -1 Then
                                            bFieldFound = True

                                            ' Get length quantity from each of the results features in the list for 
                                            ' this map layer.
                                            k = 0
                                            For k = 0 To lRefinedResultsIDsObject.Count - 1

                                                pFeature = pFeatureLayer.FeatureClass.GetFeature(lRefinedResultsIDsObject(k)._FID)
                                                vTemp = pFeature.Value(iFieldIndex)
                                                Try
                                                    dTemp = Convert.ToDouble(vTemp)
                                                Catch ex As Exception
                                                    MsgBox("Can't convert value in DistanceToMouth calculation retrieved from field to type 'double' (decimal) for feature class: " & pFeatureLayer.FeatureClass.AliasName)
                                                End Try

                                                dTotalLength = dTemp + dTotalLength
                                            Next


                                            ' For this feature / flag have found the total length now, for this feature class.  
                                            ' if there are other feature classes continue to loop through

                                        Else
                                            bFieldFound = False
                                            MsgBox("Can't find the field for this feature class, even though it's a network line... feauture class: " & pFeatureLayer.FeatureClass.AliasName)
                                        End If

                                    Catch ex As Exception
                                        MsgBox("Error encountered in DistanceToMouth during retrieval of fields or values from feature class: " & pFeatureLayer.FeatureClass.AliasName & ex.Message)

                                    End Try
                                End If
                            End If

                        Next ' edge network feature class

                        pFCID_FID_dTotal_Source = New FCID_FID_dTotalObject(selectedFCID, selectedFID, dTotalLength)
                        lFCID_FID_dTotal_Source.Add(pFCID_FID_dTotal_Source)
                        dTotalLength = 0
                    End If 'bSource is true

                    pNetworkAnalysisExtFlags.ClearFlags()
                    selectedFID = pEnumIDs.Next

                Loop ' next selected feature / flag in the layer

                If bMouth = True Then
                    ' ===== Write Total Distance to mouth to Attribute Table =====
                    ' 
                    ' update the selected features attribute table with the total distance to mouth

                    pFeatureLayer = CType(pMap.Layer(iIndex), IFeatureLayer)
                    iFieldMatch = pFeatureLayer.FeatureClass.FindField(sFieldNameMouth)
                    pEditor.StartEditing(pFeatureLayer.FeatureClass.FeatureDataset.Workspace)

                    pFeatureLayer = CType(pMap.Layer(iIndex), IFeatureLayer)
                    iFieldMatch = pFeatureLayer.FeatureClass.FindField(sFieldNameMouth)

                    j = 0
                    For j = 0 To lFCID_FID_dTotal_Mouth.Count - 1

                        Try
                            ' Get the object where FID matches
                            pFCID_FID_dTotalObject_Predicate = New FCID_FID_dTotalObject_Predicate(lFCID_FID_dTotal_Mouth(j)._FID)
                            pFCID_FID_dTotal_Mouth = lFCID_FID_dTotal_Mouth.Find(AddressOf pFCID_FID_dTotalObject_Predicate.CompareByFID)

                            pFeature = pFeatureLayer.FeatureClass.GetFeature(lFCID_FID_dTotal_Mouth(j)._FID)
                            sOIDName = pFeature.Table.OIDFieldName
                            pQueryFilter = New QueryFilterClass()
                            pQueryFilter.WhereClause = sOIDName + "= " + Convert.ToString(lFCID_FID_dTotal_Mouth(j)._FID)
                            pField = pFeatureLayer.FeatureClass.Fields.Field(iFieldMatch)

                            pQueryFilter.SubFields = (sOIDName + ", " + pField.Name)

                            pFeatureCursor = pFeatureLayer.FeatureClass.Update(pQueryFilter, False)
                            pFeature = pFeatureCursor.NextFeature
                            'pFeature = pFeatureLayer.FeatureClass.GetFeature(lFCID_FID_dTotal(j)._FID)

                            pFeature.Value(iFieldMatch) = pFCID_FID_dTotal_Mouth._dTotal
                            pFeatureCursor.UpdateFeature(pFeature)

                        Catch ex As Exception
                            MsgBox("Couldn't update featureclass field with the distance-to-mouth. " + ex.Message)
                        End Try

                    Next
                    pEditor.StopEditing(True)
                End If

                If bSource = True Then
                    ' ===== Write Total Distance to Source to Attribute Table =====
                    ' 
                    ' update the selected features attribute table with the total distance to mouth

                    pFeatureLayer = CType(pMap.Layer(iIndex), IFeatureLayer)
                    iFieldMatch = pFeatureLayer.FeatureClass.FindField(sFieldNameSource)
                    pEditor.StartEditing(pFeatureLayer.FeatureClass.FeatureDataset.Workspace)

                    pFeatureLayer = CType(pMap.Layer(iIndex), IFeatureLayer)
                    iFieldMatch = pFeatureLayer.FeatureClass.FindField(sFieldNameSource)

                    j = 0
                    For j = 0 To lFCID_FID_dTotal_Source.Count - 1

                        Try
                            ' Get the object where FID matches
                            pFCID_FID_dTotalObject_Predicate = New FCID_FID_dTotalObject_Predicate(lFCID_FID_dTotal_Source(j)._FID)
                            pFCID_FID_dTotal_Source = lFCID_FID_dTotal_Source.Find(AddressOf pFCID_FID_dTotalObject_Predicate.CompareByFID)

                            pFeature = pFeatureLayer.FeatureClass.GetFeature(lFCID_FID_dTotal_Source(j)._FID)
                            sOIDName = pFeature.Table.OIDFieldName
                            pQueryFilter = New QueryFilterClass()
                            pQueryFilter.WhereClause = sOIDName + "= " + Convert.ToString(lFCID_FID_dTotal_Source(j)._FID)
                            pField = pFeatureLayer.FeatureClass.Fields.Field(iFieldMatch)

                            pQueryFilter.SubFields = (sOIDName + ", " + pField.Name)

                            pFeatureCursor = pFeatureLayer.FeatureClass.Update(pQueryFilter, False)
                            pFeature = pFeatureCursor.NextFeature
                            'pFeature = pFeatureLayer.FeatureClass.GetFeature(lFCID_FID_dTotal(j)._FID)

                            pFeature.Value(iFieldMatch) = pFCID_FID_dTotal_Source._dTotal
                            pFeatureCursor.UpdateFeature(pFeature)

                        Catch ex As Exception
                            MsgBox("Couldn't update featureclass field with the distance-to-mouth. " + ex.Message)
                        End Try

                    Next
                    pEditor.StopEditing(True)
                End If

                ' ------------------------------------
                ' clear flags
            Next ' source layer

            ' ===============================================
            ' Restore original flags - both junction and edge
            ' Restore original barriers - both junction and edge

            ' Restore flags :

            ' ======================== RESET FLAGS ================================
            ' restore all EDGE flags
            ' Edge flags use a list of iFlagdisplay objects
            i = 0
            For i = 0 To lEdgeFlagDisplaylist.Count - 1

                '' get eid from iflagdisplay object stored in list
                'iEID = pNetElements.GetEID(lEdgeBarrierDisplaylist(i).FeatureClassID, lEdgeBarrierDisplaylist(i).FID, lEdgeBarrierDisplaylist(i).SubID, esriElementType.esriETEdge)

                '' Query the corresponding user ID's to the element ID
                'pNetElements.QueryIDs(iEID, esriElementType.esriETEdge, iFCID, iFID, iSubID)

                '' Display the flags as a EdgeFlagDisplay type
                'pFlagDisplay = New EdgeFlagDisplay
                'pSymbol = CType(pFlagSymbol, ISymbol)
                'With pFlagDisplay
                '    .FeatureClassID = iFCID
                '    .FID = iFID
                '    .Geometry = pGeometricNetwork.GeometryForEdgeEID(iEID)
                '    .Symbol = pSymbol
                'End With
                pFlagDisplay = lEdgeFlagDisplaylist(i)

                ' Add the flags to the logical network
                pEdgeFlagDisplay = CType(pFlagDisplay, IEdgeFlagDisplay)
                pNetworkAnalysisExtFlags.AddEdgeFlag(pEdgeFlagDisplay)
            Next

            ' restore all JUNCTION Flags
            i = 0
            pOriginalJuncFlagsList.Reset()
            For i = 0 To pOriginalJuncFlagsList.Count - 1

                iEID = pOriginalJuncFlagsList.Next
                ' Query the corresponding user ID's to the element ID
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
                pJunctionFlagDisplay = CType(pFlagDisplay, IJunctionFlagDisplay)
                pNetworkAnalysisExtFlags.AddJunctionFlag(pJunctionFlagDisplay)

            Next

            ' --------------------------------------------------------------
            ' need to set up a barrier symbol (flag symbol done earlier)
            ' Flag symbol code
            Dim pBarrierSymbol As ESRI.ArcGIS.Display.ISimpleMarkerSymbol
            ' New IFlagdisplay field for displaying barriers

            pRgbColor = New ESRI.ArcGIS.Display.RgbColor
            With pRgbColor
                .Red = 255
                .Green = 0
                .Blue = 0
            End With
            pSimpleMarkerSymbol = New ESRI.ArcGIS.Display.SimpleMarkerSymbol
            With pSimpleMarkerSymbol
                .Color = pRgbColor
                .Style = ESRI.ArcGIS.Display.esriSimpleMarkerStyle.esriSMSX
                .Outline = True
                .Size = 10
            End With

            pBarrierSymbol = pSimpleMarkerSymbol

            '               Reset things the way the user had them
            ' ========================= RESET Junciton  BARRIERS ===========================
            i = 0
            For i = 0 To pOriginalJuncBarriersList.Count - 1
                iEID = pOriginalJuncBarriersList.Next
                pNetElements.QueryIDs(iEID, esriElementType.esriETJunction, iFCID, iFID, iSubID)

                ' Display the barriers as a JunctionFlagDisplay type
                pFlagDisplay = New JunctionFlagDisplay
                pSymbol = CType(pBarrierSymbol, ISymbol)
                With pFlagDisplay
                    .FeatureClassID = iFCID
                    .FID = iFID
                    .Geometry = pGeometricNetwork.GeometryForJunctionEID(iEID)
                    .Symbol = pSymbol
                End With

                ' Add the flags to the logical network
                pJunctionFlagDisplay = CType(pFlagDisplay, IJunctionFlagDisplay)
                pNetworkAnalysisExtBarriers.AddJunctionBarrier(pEdgeFlagDisplay)
            Next
            ' ====================== END RESET BARRIERS ===========================

            '               Reset things the way the user had them
            ' ========================= RESET Edge BARRIERS ===========================
            i = 0
            For i = 0 To pOriginalEdgeBarriersList.Count - 1
                iEID = pOriginalEdgeBarriersList.Next
                pNetElements.QueryIDs(iEID, esriElementType.esriETEdge, iFCID, iFID, iSubID)

                ' Display the barriers as a JunctionFlagDisplay type
                pFlagDisplay = New JunctionFlagDisplay
                pSymbol = CType(pBarrierSymbol, ISymbol)
                With pFlagDisplay
                    .FeatureClassID = iFCID
                    .FID = iFID
                    .Geometry = pGeometricNetwork.GeometryForEdgeEID(iEID)
                    .Symbol = pSymbol
                End With

                ' Add the flags to the logical network
                pJunctionFlagDisplay = CType(pFlagDisplay, IJunctionFlagDisplay)
                pNetworkAnalysisExtBarriers.AddEdgeBarrier(pEdgeFlagDisplay)
            Next
            ' ====================== END RESET BARRIERS ===========================

        Catch ex As Exception
            MsgBox("Stupid Exception! " & ex.Message)
        End Try

    End Sub

    Protected Overrides Sub OnUpdate()
        ' use the extension listener to avoid constant checks to the 
        ' map network.  The extension listener will only update the boolean
        ' check on network count if there's a map change
        ' upgrade at version 10
        ' protected override void OnUpdate()
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

        ' if there's a network present and selection on network elements
        If m_pNetworkAnalysisExt.NetworkCount > 0 Then
            Dim pMxDocument As IMxDocument = CType(My.ArcMap.Application.Document, IMxDocument)

            If pMxDocument.FocusMap.SelectionCount > 0 Then

                Dim pGeometricNetwork As IGeometricNetwork = m_pNetworkAnalysisExt.CurrentNetwork
                Dim i, iSelectionCount As Integer
                Dim pFeatureLayer As IFeatureLayer
                Dim pFeatureClass As IFeatureClass
                Dim pFeatureSelection As IFeatureSelection

                Dim pMap As IMap = pMxDocument.FocusMap

                ' If there is a selection in the map
                '   For each layer in the map
                '     If the type of layer is a member of the geometric network
                '     of type SIMPLE JUNCTION
                '       For each feature class in the list of network junctions
                '         If there's a match
                '           If there are selected features in the map layer
                '             Then enable this tool

                Dim simpleJunctionFCs As IEnumFeatureClass
                Dim simpleEdgeFCs As IEnumFeatureClass
                simpleJunctionFCs = pGeometricNetwork.ClassesByType(ESRI.ArcGIS.Geodatabase.esriFeatureType.esriFTSimpleJunction)
                simpleEdgeFCs = pGeometricNetwork.ClassesByType(ESRI.ArcGIS.Geodatabase.esriFeatureType.esriFTSimpleEdge)

                For i = 0 To pMap.LayerCount - 1
                    If pMap.Layer(i).Valid = True Then
                        ' If it's a feature layer then
                        If TypeOf pMap.Layer(i) Is IFeatureLayer Then
                            pFeatureLayer = CType(pMap.Layer(i), IFeatureLayer)
                            ' If it's a junction then set the type as a simple junction
                            If pFeatureLayer.FeatureClass.FeatureType = esriFeatureType.esriFTSimpleJunction Then
                                simpleJunctionFCs.Reset()
                                pFeatureClass = simpleJunctionFCs.Next
                                ' Cycle through the network junction feature classes
                                Do Until pFeatureClass Is Nothing
                                    ' If there is a match between the map layer FC and the junctionFC
                                    If pFeatureClass Is pFeatureLayer.FeatureClass Then
                                        ' Count number of selected features in layer
                                        pFeatureSelection = CType(pFeatureLayer, IFeatureSelection)
                                        iSelectionCount = pFeatureSelection.SelectionSet.Count
                                        If iSelectionCount > 0 Then
                                            Me.Enabled = True
                                            Exit Do
                                        End If
                                    End If
                                    pFeatureClass = simpleJunctionFCs.Next
                                Loop
                            ElseIf pFeatureLayer.FeatureClass.FeatureType = esriFeatureType.esriFTSimpleEdge Then
                                simpleEdgeFCs.Reset()
                                pFeatureClass = simpleEdgeFCs.Next
                                ' Cycle through the network junction feature classes
                                Do Until pFeatureClass Is Nothing
                                    ' If there is a match between the map layer FC and the junctionFC
                                    If pFeatureClass Is pFeatureLayer.FeatureClass Then
                                        ' Count number of selected features in layer
                                        pFeatureSelection = CType(pFeatureLayer, IFeatureSelection)
                                        iSelectionCount = pFeatureSelection.SelectionSet.Count
                                        If iSelectionCount > 0 Then
                                            Me.Enabled = True
                                            Exit Do
                                        End If
                                    End If
                                    pFeatureClass = simpleEdgeFCs.Next
                                Loop
                            End If
                        End If
                    End If
                Next
            Else ' if no selection count me.enabled =false
                Me.Enabled = False

            End If ' selection count is zero
        Else
            Me.Enabled = False
        End If
    End Sub
    Private Sub ManageResultsForm()
        ProgressForm = New frmProgress_DistanceToMouth
        Using ProgressForm As New frmProgress_DistanceToMouth
            If ProgressForm.Form_Initialize(My.ArcMap.Application) Then
                ProgressForm.lblGeonetSettings.Text = m_sGeonetsettings
                ProgressForm.lblGetNetworkLines.Text = m_sGetNetworkLines
                ProgressForm.lblCurrentFeature.Text = m_iCurrentFeature
                ProgressForm.lblTotalFeatures.Text = m_iTotalFeatures
                ProgressForm.TopMost = True
                ProgressForm.BringToFront()
                ProgressForm.ShowDialog()
            End If
        End Using
        '    Private _iCurrentFeature As Integer
        '    Private _iTotalFeatures As Integer
        '    Private _sGetNetworkLines As String

        'Public Sub New(ByRef icurrentfeature As Integer, ByVal itotalfeatures As Integer, ByVal sgetnetworklines As String)
        '    Me._iCurrentFeature = icurrentfeature
        '    Me._iTotalFeatures = itotalfeatures
        '    Me._sGetNetworkLines = sgetnetworklines
    End Sub


    Private Function TraceFlowSolverSetup3() As ITraceFlowSolver
        ' Prepares the network for tracing

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
            Dim pTraceFlowSolverGEN As ITraceFlowSolverGEN
            pTraceFlowSolverGEN = CType(pNetSolver, ITraceFlowSolverGEN)
            pTraceFlowSolverGEN.PutEdgeOrigins(pEdgeFlags)

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
        TraceFlowSolverSetup3 = pTraceFlowSolver

    End Function
    Private Class FindByFCIDPredicate
        ' this class should return all objects in an object list
        ' with a given FCID
        Private _fcid As Integer

        Public Sub New(ByVal fcid As Integer)
            Me._fcid = fcid
        End Sub

        ' this is a double-check and should 
        ' return true if there's a match between
        ' BOTH the layer and EID of the sink/barrier
        Public Function CompareByFCID(ByVal obj As ResultsIDsObject) As Boolean
            Return (_fcid = obj._FCID)
        End Function
    End Class
    Private Class FindByFCIDPredicate2
        ' this class should return all objects in an object list
        ' with a given FCID
        Private _fcid As Integer

        Public Sub New(ByVal fcid As Integer)
            Me._fcid = fcid
        End Sub

        ' this is a double-check and should 
        ' return true if there's a match between
        ' BOTH the layer and EID of the sink/barrier
        Public Function CompareByFCID(ByVal obj As LayerFCIDAndQuanField) As Boolean
            Return (_fcid = obj._FCID)
        End Function
    End Class

    Private Class FCID_FID_dTotalObject_Predicate
        ' this class should return all objects in an object list
        ' with a given FCID
        Private _fid As Integer

        Public Sub New(ByVal fid As Integer)
            Me._fid = fid
        End Sub

        ' this is a double-check and should 
        ' return true if there's a match between
        ' BOTH the layer and EID of the sink/barrier
        Public Function CompareByFID(ByVal obj As FCID_FID_dTotalObject) As Boolean
            Return (_fid = obj._FID)
        End Function
    End Class

    Protected Overrides Sub Finalize()
        MyBase.Finalize()
    End Sub



End Class
