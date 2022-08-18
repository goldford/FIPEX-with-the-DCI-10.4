Imports System.Runtime.InteropServices
Imports System.Drawing
Imports ESRI.ArcGIS.ADF.BaseClasses
Imports ESRI.ArcGIS.ADF.CATIDs
Imports ESRI.ArcGIS.Framework
Imports ESRI.ArcGIS.Geometry
Imports ESRI.ArcGIS.ArcMapUI
Imports ESRI.ArcGIS.esriSystem
Imports ESRI.ArcGIS.EditorExt
Imports ESRI.ArcGIS.NetworkAnalysis
Imports ESRI.ArcGIS.Geodatabase
Imports ESRI.ArcGIS.Carto
Imports ESRI.ArcGIS.Display
Imports ESRI.ArcGIS.Geoprocessing
Imports ESRI.ArcGIS.SystemUI
Imports System.Windows.Forms

Public Class PlaceRemoveBarrier
    Inherits ESRI.ArcGIS.Desktop.AddIns.Tool
    'Private m_pGeometricNetwork As IGeometricNetwork
    Private m_UNAExt As IUtilityNetworkAnalysisExt
    Private m_pNetworkAnalysisExt As INetworkAnalysisExt
    Private m_FiPEx__1 As FishPassageExtension
    Public Sub New()

    End Sub

    Protected Overloads Overrides Sub OnUpdate()

        'On Error Resume Next

        ' use the extension listener to avoid constant checks to the 
        ' map network.  The extension listener will only update the boolean
        ' check on network count if there's a map change
        ' upgrade at version 10
        ' protected override void OnUpdate()
        If m_FiPEx__1 Is Nothing Then
            m_FiPEx__1 = FiPEX_ArcMap_10p4_up_AddIn_dotNet45_2022.FishPassageExtension.GetExtension()
        End If
        If m_UNAExt Is Nothing Then
            m_UNAExt = FiPEX_ArcMap_10p4_up_AddIn_dotNet45_2022.FishPassageExtension.GetUNAExt
        End If
        'Dim FiPEx__1 As FishPassageExtension = FiPEX_AddIn_dotNet35_2.FishPassageExtension.GetExtension
        'Dim pUNAExt As IUtilityNetworkAnalysisExt = FiPEX_AddIn_dotNet35_2.FishPassageExtension.GetUNAExt
        If m_pNetworkAnalysisExt Is Nothing Then
            m_pNetworkAnalysisExt = CType(m_UNAExt, INetworkAnalysisExt)
        End If

        If m_pNetworkAnalysisExt.NetworkCount > 0 Then
            Me.Enabled = True
        Else
            Me.Enabled = False
        End If
        'Me.Enabled = FiPEx__1.HasNetwork

        ' If this is getting updated
        If Me.Enabled = True Then
            If m_FiPEx__1.m_bBarriersLoaded = False Then

                Dim iBarrierEIDCount As Integer = m_FiPEx__1.pPropset.GetProperty("numBarrierEIDs")
                If iBarrierEIDCount > 0 Then

                    ' Get the active network
                    ' The network shouldn't be empty because the extension wouldn't be enabled otherwise
                    Dim pGeometricNetwork As IGeometricNetwork
                    pGeometricNetwork = m_pNetworkAnalysisExt.CurrentNetwork

                    ' Get the network elements for EID retrieval
                    Dim pNetwork As INetwork
                    Dim pNetElements As INetElements
                    pNetwork = pGeometricNetwork.Network
                    pNetElements = CType(pNetwork, INetElements)

                    Dim pNetworkAnalysisExtBarriers As INetworkAnalysisExtBarriers
                    pNetworkAnalysisExtBarriers = CType(m_UNAExt, INetworkAnalysisExtBarriers)

                    ' if current network has no Barriers loaded (double ensure this isn't run twice - 
                    '  only at the document load)
                    If pNetworkAnalysisExtBarriers.JunctionBarrierCount = 0 Then

                        Dim i As Integer = 0
                        Dim iEID, iFID, iFCID, iSubID As Integer
                        Dim lBarrierEIDs As New List(Of Integer)
                        For i = 0 To iBarrierEIDCount - 1
                            iEID = Convert.ToInt32(m_FiPEx__1.pPropset.GetProperty("barrierEID" + i.ToString))
                            lBarrierEIDs.Add(iEID)
                        Next

                        Dim pFlagDisplay As IFlagDisplay
                        Dim pRgbColor As IRgbColor
                        pRgbColor = New RgbColor

                        ' Set the flag symbol color and parameters
                        Dim pSimpleMarkerSymbol As ISimpleMarkerSymbol
                        ' For the Flag marker
                        ' For the Barrier marker
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

                        ' Result is a global variable containing a flag marker
                        'm_FlagSymbol4 = pSimpleMarkerSymbol
                        Dim pSymbol As ISymbol
                        Dim pJuncFlagDisplay As IJunctionFlagDisplay

                        i = 0
                        For i = 0 To lBarrierEIDs.Count - 1
                            iEID = lBarrierEIDs(i)
                            ' just a safe guard against a screwup - ieid = 0
                            If iEID <> 0 Then
                                Try
                                    pNetElements.QueryIDs(iEID, esriElementType.esriETJunction, iFCID, iFID, iSubID)

                                    ' Display the Barriers as a JunctionFlagDisplay type
                                    pFlagDisplay = New JunctionFlagDisplay
                                    pSymbol = CType(pSimpleMarkerSymbol, ISymbol)
                                    With pFlagDisplay
                                        .FeatureClassID = iFCID
                                        .FID = iFID
                                        .Geometry = pGeometricNetwork.GeometryForJunctionEID(iEID)
                                        .Symbol = pSymbol
                                    End With

                                    ' Add the flags to the logical network
                                    pJuncFlagDisplay = CType(pFlagDisplay, IJunctionFlagDisplay)
                                    pNetworkAnalysisExtBarriers.AddJunctionBarrier(pJuncFlagDisplay)
                                Catch ex As Exception
                                    MsgBox("A network barrier saved in the FIPEX could not be found anymore " + _
                                           "in this network. " + ex.Message)

                                End Try
                            End If ' ieid = 0
                        Next
                        ' flag count > 0 
                        Dim pActiveView As IActiveView = CType(My.ArcMap.Document.FocusMap, IActiveView)
                        pActiveView.Refresh()
                    End If 'there were any flags set for this network

                End If
                m_FiPEx__1.m_bBarriersLoaded = True
                m_FiPEx__1.pPropset.SetProperty("barrierEIDsLoadedyn", m_FiPEx__1.m_bBarriersLoaded)
            End If
        End If
    End Sub
    Protected Overrides Sub OnMouseDown(ByVal arg As MouseEventArgs)

        If m_FiPEx__1 Is Nothing Then
            m_FiPEx__1 = FiPEX_ArcMap_10p4_up_AddIn_dotNet45_2022.FishPassageExtension.GetExtension()
        End If
        If m_UNAExt Is Nothing Then
            m_UNAExt = FiPEX_ArcMap_10p4_up_AddIn_dotNet45_2022.FishPassageExtension.GetUNAExt
        End If
        If m_pNetworkAnalysisExt Is Nothing Then
            m_pNetworkAnalysisExt = CType(m_UNAExt, INetworkAnalysisExt)
        End If

        Dim pNetworkAnalysisExt As INetworkAnalysisExt = CType(m_UNAExt, INetworkAnalysisExt)
        Dim pGeometricNetwork As IGeometricNetwork = pNetworkAnalysisExt.CurrentNetwork
        Dim pMxDocument As IMxDocument = CType(My.ArcMap.Application.Document, IMxDocument)
        Dim pMap As IMap = pMxDocument.FocusMap
        Dim pMxApp As IMxApplication = CType(My.ArcMap.Application, IMxApplication)
        Dim pActiveView As IActiveView = CType(pMap, IActiveView)
        Dim pNetwork As INetwork = pGeometricNetwork.Network
        Dim pNetElements As INetElements = CType(pNetwork, INetElements)
        Dim pTraceTasks As ITraceTasks = CType(m_UNAExt, ITraceTasks)
        Dim pNetworkAnalysisExtResults As INetworkAnalysisExtResults = CType(m_UNAExt, INetworkAnalysisExtResults)

        ' ------------------------------------
        ' Get the point in map units
        Dim pPoint As IPoint = New ESRI.ArcGIS.Geometry.Point
        pPoint = pMxApp.Display.DisplayTransformation.ToMapPoint(arg.X, arg.Y)
        ' find the nearest junction element to this Point
        Dim xEID As Integer = 0
        Dim outPoint As IPoint = New ESRI.ArcGIS.Geometry.Point
        Dim pointToEID As IPointToEID = New PointToEID
        With pointToEID
            .GeometricNetwork = pGeometricNetwork
            .SourceMap = pMap
            .SnapTolerance = 100     ' set a snap tolerance of 100 map units
            .GetNearestJunction(pPoint, xEID, outPoint)
        End With

        ' Exit if no point is found
        If xEID = 0 Then
            Exit Sub
        End If

        ' ------------------------------------ 
        ' Get the current network Barriers
        Dim pNetworkAnalysisExtBarriers As INetworkAnalysisExtBarriers = CType(m_UNAExt, INetworkAnalysisExtBarriers)

        Dim pFlagDisplay As IFlagDisplay
        Dim pOriginaljuncBarriersListGEN As IEnumNetEIDBuilderGEN = New EnumNetEIDArray
        Dim iEID As Integer
        Dim bBarrierCheck As Boolean = False
        Dim i As Integer = 0
        ' If there ARE Barriers
        If pNetworkAnalysisExtBarriers.JunctionBarrierCount <> 0 Then
            ' Get the EIDs
            For i = 0 To pNetworkAnalysisExtBarriers.JunctionBarrierCount - 1
                ' Use the bFlagDisplay to retrieve the EIDs of the junction Barriers
                pFlagDisplay = CType(pNetworkAnalysisExtBarriers.JunctionBarrier(i), IFlagDisplay)
                iEID = pNetElements.GetEID(pFlagDisplay.FeatureClassID, pFlagDisplay.FID, pFlagDisplay.SubID, esriElementType.esriETJunction)
                If iEID = xEID Then
                    bBarrierCheck = True
                End If
                pOriginaljuncBarriersListGEN.Add(iEID)
            Next
        End If

        ' QI to and get an array interface that has 'count' and 'next' methods
        Dim pOriginaljuncBarriersList As IEnumNetEID = CType(pOriginaljuncBarriersListGEN, IEnumNetEID)

        ' ------------------------------------
        ' Add or Remove a Barrier

        ' Set up the Barrier marker symbol object
        Dim pSimpleMarkerSymbol As ISimpleMarkerSymbol
        Dim pRgbColor As IRgbColor = New RgbColor

        ' For the Barrier marker
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

        Dim pJunctionFlagDisplay As IJunctionFlagDisplay
        Dim netElements As INetElements = CType(pGeometricNetwork.Network, INetElements)
        Dim FCID As Integer, FID As Integer, subID As Integer

        ' convert the EID to a feature class ID, feature ID, and sub ID            
        netElements.QueryIDs(xEID, esriElementType.esriETJunction, FCID, FID, subID)

        ' if there is no Barrier on the junction add one
        If bBarrierCheck = False Then

            ' create a new JunctionFlagDisplay object and populate it
            pJunctionFlagDisplay = New JunctionFlagDisplay
            pFlagDisplay = CType(pJunctionFlagDisplay, IFlagDisplay)
            With pFlagDisplay
                .FeatureClassID = FCID
                .FID = FID
                .SubID = subID
                .Geometry = pGeometricNetwork.GeometryForJunctionEID(xEID)
                .Symbol = CType(pSimpleMarkerSymbol, ISymbol)
            End With

            ' add the JunctionFlagDisplay object to the Network Analysis extension
            pJunctionFlagDisplay = CType(pFlagDisplay, IJunctionFlagDisplay)
            pNetworkAnalysisExtBarriers.AddJunctionBarrier(pJunctionFlagDisplay)

            ' label the new barrier if needed
            Call LabelBarrier(FCID, FID)

            ' If there is already a Barrier present then remove it
        ElseIf bBarrierCheck = True Then

            ' label the new barrier if needed
            Call UnLabelBarrier(FCID, FID)

            ' Cannot just remove a junction Barrier
            ' -needs to be two-step process
            pNetworkAnalysisExtBarriers.ClearBarriers()

            ' Filter the EIDs and make a new list without the clicked junction
            Dim pNewJuncBarriersListGEN As IEnumNetEIDBuilderGEN = New EnumNetEIDArray
            pOriginaljuncBarriersList.Reset()
            For i = 0 To pOriginaljuncBarriersList.Count - 1
                iEID = pOriginaljuncBarriersList.Next
                If iEID <> xEID Then
                    pNewJuncBarriersListGEN.Add(iEID)
                End If
            Next
            Dim pNewJuncBarriersList As IEnumNetEID = CType(pNewJuncBarriersListGEN, IEnumNetEID)

            pNewJuncBarriersList.Reset()
            For i = 0 To pNewJuncBarriersList.Count - 1
                Try

                    iEID = pNewJuncBarriersList.Next
                    netElements.QueryIDs(iEID, esriElementType.esriETJunction, FCID, FID, subID)
                    ' create a new JunctionFlagDisplay object and populate it
                    pJunctionFlagDisplay = New JunctionFlagDisplay
                    pFlagDisplay = CType(pJunctionFlagDisplay, IFlagDisplay)
                    With pFlagDisplay
                        .FeatureClassID = FCID
                        .FID = FID
                        .SubID = subID
                        .Geometry = pGeometricNetwork.GeometryForJunctionEID(iEID)
                        .Symbol = CType(pSimpleMarkerSymbol, ISymbol)
                    End With

                    ' add the JunctionFlagDisplay object to the Network Analysis extension
                    pJunctionFlagDisplay = CType(pFlagDisplay, IJunctionFlagDisplay)
                    pNetworkAnalysisExtBarriers.AddJunctionBarrier(pJunctionFlagDisplay)
                Catch ex As Exception
                    MsgBox("A network Flag element saved in FIPEX can no longer be found in this network. " + _
                           ex.Message)

                End Try
            Next
        End If

        pActiveView.Refresh()
    End Sub
    Private Sub LabelBarrier(ByVal iFCID As Integer, ByVal iFID As Integer)

        ' Created By: Greig Oldford
        ' Date: July 5, 2009
        ' Purpose: Label barriers or flags if no label is present
        '          using user-set field from extension settings
        '   
        ' 1.0 Read the extension settings
        ' 2.0 label the flag if needed 
        ' 
        ' Bug Note: Since there are issues with finding visible label
        '           elements (see notes further down) the workaround
        '           shows labels as visible if they have been turned on
        '           and then off in ArcMap.  So if you have turned labels
        '           off they might still be in the annotation properties
        '           and return as 'visible.'

        Dim pMxDocument As IMxDocument = CType(My.ArcMap.Document, IMxDocument)
        Dim pMap As IMap = pMxDocument.FocusMap
        Dim pMxApp As IMxApplication = CType(My.ArcMap.ThisApplication, IMxApplication)
        If m_FiPEx__1 Is Nothing Then
            m_FiPEx__1 = FiPEX_ArcMap_10p4_up_AddIn_dotNet45_2022.FishPassageExtension.GetExtension()
        End If


        ' ------------------------------------
        ' 1.0 Read Extension Barrier label settings
        Dim lBarrierIDs As List(Of BarrierIDObj) = New List(Of BarrierIDObj)
        Dim pBarrierIDObj As New BarrierIDObj(Nothing, Nothing, Nothing, Nothing, Nothing)

        Dim j, i As Integer
        Dim iBarrierIDs As Integer
        Dim sBarrierIDLayer, sBarrierIDField As String

        ' If settings have been set by the user then load them
        If m_FiPEx__1.m_bLoaded = True Then
            ' Get the barrier ID Fields
            iBarrierIDs = Convert.ToInt32(m_FiPEx__1.pPropset.GetProperty("numBarrierIDs"))
            If iBarrierIDs > 0 Then
                For j = 0 To iBarrierIDs - 1
                    sBarrierIDLayer = Convert.ToString(m_FiPEx__1.pPropset.GetProperty("BarrierIDLayer" + j.ToString))
                    sBarrierIDField = Convert.ToString(m_FiPEx__1.pPropset.GetProperty("BarrierIDField" + j.ToString))
                    ' Object to retrieve barrier label field (do not need other obj. params so set to 'nothing')
                    pBarrierIDObj = New BarrierIDObj(sBarrierIDLayer, sBarrierIDField, Nothing, Nothing, Nothing)
                    lBarrierIDs.Add(pBarrierIDObj)
                Next
            End If
        End If

        ' ------------------------------------
        ' 2.0 Label the flag if needed

        Dim pAnnoLayerPropsColl As IAnnotateLayerPropertiesCollection
        'Dim pMapAnnoLayerPropsColl As IAnnotateLayerPropertiesCollection
        Dim pAnnoLayerProps As IAnnotateLayerProperties
        Dim pNewLELayerProps As ILabelEngineLayerProperties
        Dim aLELayerProps As ILabelEngineLayerProperties
        Dim pEnumVisibleElements As IElementCollection
        'Dim pElement As IElement
        'Dim pTextElement As ITextElement
        'Dim pAnnotateMapProps As IAnnotateMapProperties
        'Dim pAnnotateMap As IAnnotateMap
        Dim pEnumInVisibleElements As IElementCollection
        Dim sLabelField, sLabelValue As String
        Dim sSQL, strOIDName As String
        Dim sPreviousWhereClause As String = ""
        Dim pGFLayer As IGeoFeatureLayer
        Dim pFLayer As IFeatureLayer
        'Dim sText As String
        Dim bLabelMatch As Boolean = False
        Dim bUserLabel As Boolean = False
        Dim bString As Boolean = False
        Dim pFieldType As esriFieldType
        Dim iFieldIndex As Integer
        Dim pFields As IFields
        Dim bClassMatch As Boolean
        Dim iClassNum As Integer
        Dim sSearchString As String
        Dim pFeature As IFeature
        Dim pFeatureClass As IFeatureClass

        ' For each layer in the map
        '   If it's a feature layer
        '     If it's the same layer as the new flag is being placed on
        '       For each layer in the BarrierID list
        '         If it matches this layer
        '           If the field is found in this layer
        '             Then use this field as a label        '
        '       Get the annotation properties of the layer
        '       For each of the properties
        '         Get the visible elements
        '         If any one label matches the ObjectID of the element
        '           Alert that there's a match (so labels won't be changed)
        '         If there was no match found then
        '           Create a new label class called "FlagID"



        For i = 0 To pMap.LayerCount - 1
            If pMap.Layer(i).Valid = True Then
                If TypeOf pMap.Layer(i) Is IFeatureLayer Then
                    pFLayer = CType(pMap.Layer(i), IFeatureLayer)
                    If pFLayer.FeatureClass.FeatureClassID = iFCID Then

                        bLabelMatch = False 'reset match variable for label
                        pGFLayer = CType(pMap.Layer(i), IGeoFeatureLayer)

                        ' For each layer in the BarrierID list
                        '   If it matches this layer
                        '     If the field is found in this layer
                        '       Then use this field as a label
                        bUserLabel = False
                        If lBarrierIDs IsNot Nothing Then
                            If lBarrierIDs.Count <> 0 Then
                                For j = 0 To lBarrierIDs.Count - 1
                                    If lBarrierIDs.Item(j).Layer = pFLayer.Name Then

                                        pFeatureClass = pFLayer.FeatureClass
                                        pFeature = pFeatureClass.GetFeature(iFID)
                                        pFields = pFeature.Fields

                                        If pFields.FindField(lBarrierIDs.Item(j).Field) <> -1 Then

                                            sLabelField = lBarrierIDs.Item(j).Field
                                            Try

                                                sLabelValue = Convert.ToString(pFeature.Value(pFields.FindField(sLabelField)))
                                            Catch ex As Exception
                                                MsgBox("Could not convert label value " + sLabelValue _
                                                       + " to type 'string' for feature in feature class " + pFeatureClass.AliasName _
                                                       + ex.Message)
                                                Exit For

                                            End Try

                                            ' If a value was returned then set the alert variable true
                                            If sLabelValue <> "" Then

                                                bUserLabel = True

                                                ' Get the field type of the field because if it is a string
                                                ' the sql requires quotation wrappers. 
                                                iFieldIndex = pFeatureClass.FindField(sLabelField)
                                                pFieldType = pFields.Field(iFieldIndex).Type

                                                If pFieldType = esriFieldType.esriFieldTypeString Then
                                                    bString = True
                                                End If

                                            End If
                                        End If
                                    End If
                                Next
                            End If
                        End If

                        ' Since pAnnotationLayerPropertiesCollection.QueryItem does not return the proper
                        ' visible elements a workaround found here 
                        ' http://forums.esri.com/Thread.asp?c=93&f=993&t=164358 
                        ' or here for details
                        ' http://edn.esri.com/index.cfm?fa=codeExch.sampleDetail&pg=/arcobjects/9.1/Samples/Cartography/Labeling_and_Annotation/LabelsToMapAnno.htm
                        ' was attempted - cloning the annotation properties did not work, though. 


                        'Dim propsIndex As Integer
                        'Dim pClone As IClone

                        '' Clone the map annotation properties collection
                        'For propsIndex = 0 To (pAnnoLayerPropsColl.Count - 1)
                        '    pAnnoLayerPropsColl.QueryItem(propsIndex, pAnnoLayerProps1, Nothing, Nothing)
                        '    If Not pAnnoLayerProps1 Is Nothing Then
                        '        'Clone the properties and add them to the new collection
                        '        pClone = CType(pAnnoLayerProps1, IClone)
                        '        pMapAnnoLayerPropsColl.Add(CType(pClone.Clone, IAnnotateLayerProperties))
                        '    End If
                        'Next

                        'pEnumVisibleElements = New ElementCollection
                        'pEnumInVisibleElements = New ElementCollection
                        'pMapAnnoLayerPropsColl.QueryItem(0, pAnnoLayerProps, pEnumVisibleElements, pEnumInVisibleElements)

                        ' See http://forums.esri.com/Thread.asp?c=93&f=992&t=69162#180195 for 
                        ' discussion of this code
                        ' Get the Annotation Collection
                        pAnnoLayerPropsColl = pGFLayer.AnnotationProperties

                        Dim propsIndex As Integer
                        For propsIndex = 0 To (pAnnoLayerPropsColl.Count - 1)

                            pEnumVisibleElements = New ElementCollection
                            pEnumInVisibleElements = New ElementCollection
                            pAnnoLayerPropsColl.QueryItem(propsIndex, pAnnoLayerProps, pEnumVisibleElements, pEnumInVisibleElements)

                            ' If there is already a class called "BarrierOrFlagID" 
                            ' Then get the where clause and save it for later
                            If pAnnoLayerProps.Class.ToString = "BarrierOrFlagID" Then
                                sPreviousWhereClause = pAnnoLayerProps.WhereClause.ToString
                                bClassMatch = True
                                iClassNum = propsIndex
                            End If

                            ' If the layer had no labels already then make sure the new 
                            ' class (default) is defaulted to 'off' - don't show labels
                            ' otherwise it will show labels for the 'default' class. 
                            If propsIndex = 0 And pAnnoLayerPropsColl.Count = 1 Then
                                If pGFLayer.DisplayAnnotation = False Then
                                    pAnnoLayerProps.DisplayAnnotation = False
                                End If
                            End If
                        Next ' PropsIndex

                        ' TEMP SOLUTION
                        ' In where clause from the class, if found, search for the current feature
                        '   Check if using user settings for label to create search string
                        If bClassMatch = True And sPreviousWhereClause <> "" And bUserLabel = True Then
                            If bString = True Then
                                sSearchString = sLabelField + " = '" + sLabelValue + "'"
                            Else
                                sSearchString = sLabelField + " = " + sLabelValue
                            End If
                        Else
                            sSearchString = pGFLayer.FeatureClass.OIDFieldName & " = " & iFID.ToString
                        End If
                        If InStr(sPreviousWhereClause, sSearchString) <> 0 Then
                            bLabelMatch = True
                        End If

                        ' if there are duplicate label names in column
                        ' do not do anything if feature label is already there

                        If sPreviousWhereClause = "" Then
                            bLabelMatch = False
                        ElseIf bClassMatch = False Then
                            bLabelMatch = False
                        ElseIf bUserLabel = True Then
                            ' check if search string is there
                            If bString = True Then
                                sSearchString = sLabelField + " = '" + sLabelValue + "'"
                            Else
                                sSearchString = sLabelField + " = " + sLabelValue
                            End If
                            If InStr(sPreviousWhereClause, sSearchString) <> 0 Then
                                bLabelMatch = True
                            End If
                        End If

                        ' If there was no match between the visible labels and the feature's
                        ' value in the label field
                        If bLabelMatch = False Then

                            ' If we're using a label field specified by the user 
                            ' then use that field to label, otherwise use the FID field
                            If sPreviousWhereClause <> "" Then
                                If bUserLabel = True Then
                                    If bString = True Then
                                        sSQL = sPreviousWhereClause & " OR " & sLabelField & " = " & "'" & sLabelValue & "'"
                                    Else
                                        sSQL = sPreviousWhereClause & " OR " & sLabelField & " = " & sLabelValue
                                    End If
                                Else
                                    strOIDName = pGFLayer.FeatureClass.OIDFieldName
                                    sSQL = sPreviousWhereClause & " OR " & strOIDName & " = " & iFID
                                    sLabelField = strOIDName
                                End If
                            Else
                                If bUserLabel = True Then
                                    If bString = True Then
                                        sSQL = sLabelField & " = " & "'" & sLabelValue & "'"
                                    Else
                                        sSQL = sLabelField & " = " & sLabelValue
                                    End If
                                Else
                                    strOIDName = pGFLayer.FeatureClass.OIDFieldName
                                    sSQL = strOIDName & " = " & iFID
                                    sLabelField = strOIDName
                                End If
                            End If

                            ' If there was a class match found
                            '   Then get that annotation layer properties set
                            '    set the 'where' clause
                            ' If there was no match found
                            '    need to add a class so name it
                            '    set the 'where' clause
                            ' See here for explanation of new layerprops class
                            ' http://resources.esri.com/help/9.3/arcgisengine/dotnet/d3f93845-fedc-42f1-827b-912038c6271b.htm
                            Try
                                If bClassMatch = True Then
                                    pAnnoLayerPropsColl.QueryItem(iClassNum, pAnnoLayerProps, Nothing, Nothing)
                                    pAnnoLayerProps.WhereClause = sSQL
                                    pAnnoLayerProps.DisplayAnnotation = True
                                Else
                                    pNewLELayerProps = New LabelEngineLayerPropertiesClass()
                                    pAnnoLayerProps = CType(pNewLELayerProps, IAnnotateLayerProperties)
                                    pAnnoLayerProps.Class = "BarrierOrFlagID"
                                    pAnnoLayerProps.WhereClause = sSQL
                                    pAnnoLayerProps.DisplayAnnotation = True
                                    pAnnoLayerPropsColl.Add(pAnnoLayerProps)
                                End If
                            Catch ex As Exception
                                MsgBox("An error was encountered trying to label this barrier. " + ex.Message)
                            End Try


                            aLELayerProps = CType(pAnnoLayerProps, ILabelEngineLayerProperties)
                            aLELayerProps.Expression = "[" & sLabelField & "]"

                            ' Change symbol - add halo
                            Dim pBOLP As IBasicOverposterLayerProperties = New BasicOverposterLayerProperties
                            pBOLP.PointPlacementOnTop = False
                            pBOLP.PointPlacementMethod = esriOverposterPointPlacementMethod.esriAroundPoint
                            pBOLP.LabelWeight = esriBasicOverposterWeight.esriHighWeight

                            Dim pTxtSym As ITextSymbol = New ESRI.ArcGIS.Display.TextSymbol
                            pTxtSym = aLELayerProps.Symbol

                            'mask color
                            Dim pRgbColor As IRgbColor = New RgbColor
                            ' Set the barrier symbol color and parameters
                            With pRgbColor
                                .Red = 255
                                .Green = 255
                                .Blue = 255
                            End With

                            Dim pFillSymbol As IFillSymbol = New SimpleFillSymbol
                            pFillSymbol.Color = pRgbColor

                            Dim pLineSymbol As ILineSymbol = New SimpleLineSymbol
                            pLineSymbol.Color = pRgbColor
                            pFillSymbol.Outline = pLineSymbol

                            Dim pMask As IMask
                            pMask = CType(pTxtSym, IMask)

                            pMask.MaskStyle = esriMaskStyle.esriMSHalo
                            pMask.MaskSize = 1.8
                            pMask.MaskSymbol = pFillSymbol

                            pTxtSym = CType(pMask, ITextSymbol)

                            aLELayerProps.Symbol = pTxtSym
                            aLELayerProps.BasicOverposterLayerProperties = pBOLP

                            pGFLayer.DisplayAnnotation = True
                        End If
                    End If
                End If
            End If
        Next
    End Sub

    Private Sub UnLabelBarrier(ByVal iFCID As Integer, ByVal iFID As Integer)

        ' Created By: Greig Oldford
        ' Date: July 5, 2009
        ' Purpose: Label barriers or flags if no label is present
        '          using user-set field from extension settings
        '   
        ' 1.0 Read the extension settings
        ' 2.0 label the flag if needed 
        ' 
        ' Bug Note: Since there are issues with finding visible label
        '           elements (see notes further down) the workaround
        '           shows labels as visible if they have been turned on
        '           and then off in ArcMap.  So if you have turned labels
        '           off they might still be in the annotation properties
        '           and return as 'visible.'

        If m_FiPEx__1 Is Nothing Then
            m_FiPEx__1 = FiPEX_ArcMap_10p4_up_AddIn_dotNet45_2022.FishPassageExtension.GetExtension()
        End If

        Dim pMxDocument As IMxDocument = CType(My.ArcMap.Document, IMxDocument)
        Dim pMap As IMap = pMxDocument.FocusMap
        Dim pMxApp As IMxApplication = CType(My.ArcMap.Application, IMxApplication)
        ' ------------------------------------
        ' 1.0 Read Extension Barrier label settings
        Dim lBarrierIDs As List(Of BarrierIDObj) = New List(Of BarrierIDObj)
        Dim pBarrierIDObj As New BarrierIDObj(Nothing, Nothing, Nothing, Nothing, Nothing)

        Dim j, i As Integer
        Dim iBarrierIDs As Integer
        Dim sBarrierIDLayer, sBarrierIDField As String

        ' If settings have been set by the user then load them
        If m_FiPEx__1.m_bLoaded = True Then
            ' Get the barrier ID Fields
            iBarrierIDs = Convert.ToInt32(m_FiPEx__1.pPropset.GetProperty("numBarrierIDs"))
            If iBarrierIDs > 0 Then
                For j = 0 To iBarrierIDs - 1
                    sBarrierIDLayer = Convert.ToString(m_FiPEx__1.pPropset.GetProperty("BarrierIDLayer" + j.ToString))
                    sBarrierIDField = Convert.ToString(m_FiPEx__1.pPropset.GetProperty("BarrierIDField" + j.ToString))
                    ' Object to retrieve barrier label field (do not need other obj. params so set to 'nothing')
                    pBarrierIDObj = New BarrierIDObj(sBarrierIDLayer, sBarrierIDField, Nothing, Nothing, Nothing)
                    lBarrierIDs.Add(pBarrierIDObj)
                Next
            End If
        End If

        ' ------------------------------------
        ' 2.0 Label the flag if needed

        Dim pAnnoLayerPropsColl As IAnnotateLayerPropertiesCollection
        Dim pAnnoLayerProps As IAnnotateLayerProperties
        Dim pEnumVisibleElements As IElementCollection
        Dim pEnumInVisibleElements As IElementCollection
        Dim sLabelField, sLabelValue As String
        Dim sSQL As String
        Dim sPreviousWhereClause As String = ""
        Dim pGFLayer As IGeoFeatureLayer
        Dim pFLayer As IFeatureLayer
        Dim bLabelMatch As Boolean = False
        Dim bUserLabel As Boolean = False
        Dim bString As Boolean = False
        Dim pFieldType As esriFieldType
        Dim iFieldIndex As Integer
        Dim pFields As IFields
        Dim bClassMatch As Boolean
        Dim iClassNum As Integer
        Dim sSearchString As String
        Dim pFeature As IFeature
        Dim pFeatureClass As IFeatureClass
        Dim sORSearchString, sSearchStringOR As String
        Dim iStringPosit As Integer
        ' For each layer in the map
        '   If it's a feature layer
        '     If it's the same layer as the new flag is being placed on
        '       For each layer in the BarrierID list
        '         If it matches this layer
        '           If the field is found in this layer
        '             Then use this field as a label        '
        '       Get the annotation properties of the layer
        '       For each of the properties
        '         Get the visible elements
        '         If any one label matches the ObjectID of the element
        '           Alert that there's a match (so labels won't be changed)
        '         If there was no match found then
        '           Create a new label class called "FlagID"



        For i = 0 To pMap.LayerCount - 1
            If pMap.Layer(i).Valid = True Then
                If TypeOf pMap.Layer(i) Is IFeatureLayer Then
                    pFLayer = CType(pMap.Layer(i), IFeatureLayer)
                    If pFLayer.FeatureClass.FeatureClassID = iFCID Then

                        bLabelMatch = False 'reset match variable for label
                        pGFLayer = CType(pMap.Layer(i), IGeoFeatureLayer)

                        ' For each layer in the BarrierID list
                        '   If it matches this layer
                        '     If the field is found in this layer
                        '       Then use this field as a label
                        bUserLabel = False
                        If lBarrierIDs IsNot Nothing Then
                            If lBarrierIDs.Count <> 0 Then
                                For j = 0 To lBarrierIDs.Count - 1
                                    If lBarrierIDs.Item(j).Layer = pFLayer.Name Then

                                        pFeatureClass = pFLayer.FeatureClass
                                        pFeature = pFeatureClass.GetFeature(iFID)
                                        pFields = pFeature.Fields

                                        If pFields.FindField(lBarrierIDs.Item(j).Field) <> -1 Then

                                            sLabelField = lBarrierIDs.Item(j).Field
                                            sLabelValue = Convert.ToString(pFeature.Value(pFields.FindField(sLabelField)))

                                            ' If a value was returned then set the alert variable true
                                            If sLabelValue <> "" Then

                                                bUserLabel = True

                                                ' Get the field type of the field because if it is a string
                                                ' the sql requires quotation wrappers. 
                                                iFieldIndex = pFeatureClass.FindField(sLabelField)
                                                pFieldType = pFields.Field(iFieldIndex).Type

                                                If pFieldType = esriFieldType.esriFieldTypeString Then
                                                    bString = True
                                                End If

                                            End If
                                        End If
                                    End If
                                Next
                            End If
                        End If


                        ' See http://forums.esri.com/Thread.asp?c=93&f=992&t=69162#180195 for 
                        ' discussion of this code
                        ' Get the Annotation Collection
                        pAnnoLayerPropsColl = pGFLayer.AnnotationProperties

                        Dim propsIndex As Integer
                        For propsIndex = 0 To (pAnnoLayerPropsColl.Count - 1)

                            pEnumVisibleElements = New ElementCollection
                            pEnumInVisibleElements = New ElementCollection
                            pAnnoLayerPropsColl.QueryItem(propsIndex, pAnnoLayerProps, pEnumVisibleElements, pEnumInVisibleElements)

                            ' If there is already a class called "BarrierOrFlagID" 
                            ' Then get the where clause and save it for later
                            If pAnnoLayerProps.Class.ToString = "BarrierOrFlagID" Then
                                sPreviousWhereClause = pAnnoLayerProps.WhereClause.ToString
                                bClassMatch = True
                                iClassNum = propsIndex
                            End If


                        Next ' PropsIndex

                        ' TEMP SOLUTION
                        ' In where clause from the class, if found, search for the current feature
                        '   Check if using user settings for label to create search string
                        If bClassMatch = True And sPreviousWhereClause <> "" And bUserLabel = True Then
                            If bString = True Then
                                sSearchString = sLabelField + " = '" + sLabelValue + "' "
                            Else
                                sSearchString = sLabelField + " = " + sLabelValue + " "
                            End If
                        Else
                            sSearchString = pGFLayer.FeatureClass.OIDFieldName & " = " & iFID.ToString & " "
                        End If
                        iStringPosit = InStr(sPreviousWhereClause, sSearchString)
                        If iStringPosit <> 0 Then
                            bLabelMatch = True

                            ' need to re-search to check if this clause is found in the middle of the 
                            ' where filter string or on the end (if middle then there is an "OR" to remove)


                            ' the logic is this:
                            ' if there is only one item in the label string there will 
                            ' be no OR's found. 
                            ' If there are multiple labelled features then there will be at least one OR found.  
                            ' If this is true then there are three conditions:
                            '           1) searchstring is the only one present (no "or" before or after)
                            '           2) searchstring is at the end (preceded by "or")
                            '           3) searchstring is at the beginning (followed by "or")
                            '
                            ' If searchString is at the beginning we need to remove the trailing OR
                            ' If searchString is at the end we need to remove the preceding OR
                            ' If searchstring is in the middle we need to remove either/or trailing or preceding OR
                            '       (but must get spacing exactly right and remove trailing or preceding space)

                            Dim iMiddle, iBegin, iEnd As Boolean
                            Dim sORSearchStringOR As String
                            ' THE sSEARCHSTRING HAS A TRAILING SPACE
                            sORSearchStringOR = "OR " & sSearchString & "OR "
                            sORSearchString = "OR " & sSearchString
                            sSearchStringOR = sSearchString & "OR "

                            Try
                                iMiddle = InStr(sPreviousWhereClause, sORSearchStringOR)
                            Catch ex As Exception
                                MsgBox("Issue in Unlabel sub. Code jjsk3. " + ex.Message)
                                Exit For
                            End Try

                            'If the sSearchstring is found in the middle then we can remove it using
                            ' either the trailing or preceding OR string (sORSearchString or sSearchStringOR)
                            If iMiddle <> 0 Then

                                ' remove the searchstring from the middle of the SQL statement
                                ' but add 4 to the length os it catches the trailing OR and space
                                ' (remember the iStringPosit is not a zero-based index so it must be adjusted)
                                sSQL = sPreviousWhereClause.Remove(iStringPosit - 1, sSearchString.Length + 3)
                                pAnnoLayerPropsColl.QueryItem(iClassNum, pAnnoLayerProps, Nothing, Nothing)
                                pAnnoLayerProps.WhereClause = sSQL

                            Else
                                ' if the middle searchstring returns no results then 
                                ' test if the searchstring is in the beginning by searching
                                ' with a trailing OR. 
                                ' If it doesn't then turn up, search for a preceding OR
                                Try
                                    iBegin = InStr(sPreviousWhereClause, sSearchStringOR)
                                Catch ex As Exception
                                    MsgBox("Issue in Unlabel sub. Code dsf44. " + ex.Message)
                                    Exit For
                                End Try

                                ' if the searchstring is found in the beginning then we can remove
                                ' it using the trailing OR searchstring (sSearchStringOR)
                                If iBegin <> 0 Then
                                    ' This string is the same as the removal string for the 'middle' 
                                    ' (istringposit must be adjusted to a zero-based index, sSearchstringlength
                                    ' must be adjusted to catch the trailing OR)
                                    sSQL = sPreviousWhereClause.Remove(iStringPosit - 1, sSearchString.Length + 3)
                                    pAnnoLayerPropsColl.QueryItem(iClassNum, pAnnoLayerProps, Nothing, Nothing)
                                    pAnnoLayerProps.WhereClause = sSQL

                                Else

                                    ' If it's found at the end of theSQL statement then 
                                    ' adjust the start positiong to remove the OR preceding the 
                                    ' sSearchstring (and to make it a zero-based index)
                                    Try
                                        iEnd = InStr(sPreviousWhereClause, sSearchStringOR)
                                    Catch ex As Exception
                                        MsgBox("Issue in Unlabel sub. Code sjjo4. " + ex.Message)
                                        Exit For
                                    End Try

                                    If iEnd <> 0 Then

                                        sSQL = sPreviousWhereClause.Remove(iStringPosit - 4, sSearchString.Length)
                                        pAnnoLayerPropsColl.QueryItem(iClassNum, pAnnoLayerProps, Nothing, Nothing)
                                        pAnnoLayerProps.WhereClause = sSQL
                                    Else
                                        'warning!  can't be found, though can be found

                                    End If
                                End If

                            End If

                            '' need to re-search to check if this clause is found in the middle of the 
                            '' where filter string or on the end (if middle then there is an "OR" to remove)
                            '' Three situations: 1) searchstring is the only one present (no "or")
                            ''                   2) searchstring is at the end (preceded by "or")
                            ''                   3) searchstring is at the beginning (followed by "or")


                            'If InStr(sPreviousWhereClause, sSearchStringOR) <> 0 Then

                            '    sSearchString = sORSearchString
                            '    sSQL = sPreviousWhereClause.Remove(iStringPosit - 1, sSearchString.Length + 1)
                            '    pAnnoLayerPropsColl.QueryItem(iClassNum, pAnnoLayerProps, Nothing, Nothing)
                            '    pAnnoLayerProps.WhereClause = sSQL

                            'ElseIf InStr(sPreviousWhereClause, sORSearchString) <> 0 Then

                            '    sSearchString = "OR " + sSearchString
                            '    sSQL = sPreviousWhereClause.Remove(iStringPosit - 5, sSearchString.Length + 1)
                            '    pAnnoLayerPropsColl.QueryItem(iClassNum, pAnnoLayerProps, Nothing, Nothing)
                            '    pAnnoLayerProps.WhereClause = sSQL

                            '    ' otherwise remove the preceding "OR" from this
                            'Else
                            '    sSQL = ""
                            '    pAnnoLayerPropsColl.QueryItem(iClassNum, pAnnoLayerProps, Nothing, Nothing)
                            '    pAnnoLayerProps.WhereClause = sSQL
                            '    pAnnoLayerProps.DisplayAnnotation = False

                            'End If
                        End If
                        pGFLayer.DisplayAnnotation = True
                    End If
                End If
            End If

        Next
    End Sub

   

End Class
