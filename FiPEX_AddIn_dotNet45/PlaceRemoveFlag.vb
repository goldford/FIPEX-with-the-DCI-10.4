
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


Public Class PlaceRemoveFlag
    Inherits ESRI.ArcGIS.Desktop.AddIns.Tool
    ' Name these with 4 suffix to avoid conflict with other global variables in this project
    'Private m_application4 As IApplication

    Private m_isMouseDown As Boolean = False

    ' Private m_UtilityNetworkAnalysisExt4 As IUtilityNetworkAnalysisExt
    'Private m_pGeometricNetwork As IGeometricNetwork
    Private m_UNAExt As IUtilityNetworkAnalysisExt
    Private m_pNetworkAnalysisExt As INetworkAnalysisExt
    Private m_FiPEx__1 As FishPassageExtension
    'Private m_FlagSymbol4 As ISimpleMarkerSymbol
    'Private m_DFOExt4 As FiPEX_AddIn_dotNet35_2.FishPassageExtension


    Public Sub New()
        'MyBase.New()

        'MyBase.m_category = "FiPEx"  'localizable text 
        'MyBase.m_caption = "Places or Removes a junction flag"   'localizable text 
        'MyBase.m_message = "Places or Removes a junction flag"   'localizable text 
        'MyBase.m_toolTip = "Places or Removes a junction flag" 'localizable text 
        'MyBase.m_name = "FiPEx_PlaceRemoveFlag"  'unique id, non-localizable (e.g. "MyCategory_ArcMapTool")

        'Try
        '    'TODO: change resource name if necessary
        '    'Dim bitmapResourceName As String = Me.GetType().Name + ".bmp"
        '    'MyBase.m_bitmap = New Bitmap(Me.GetType(), bitmapResourceName)
        '    'MyBase.m_cursor = New System.Windows.Forms.Cursor(Me.GetType(), Me.GetType().Name + ".cur")
        'Catch ex As Exception
        '    System.Diagnostics.Trace.WriteLine(ex.Message, "Error in Sub New() of PlaceRemoveFlag Tool")
        'End Try
    End Sub
    Protected Overloads Overrides Sub OnUpdate()

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

        If m_pNetworkAnalysisExt.NetworkCount > 0 Then
            Me.Enabled = True
        Else
            Me.Enabled = False
        End If
        'Me.Enabled = FiPEx__1.HasNetwork

        ' If this is getting updated
        If Me.Enabled = True Then
            If m_FiPEx__1.m_bFlagsLoaded = False Then

                Dim iFlagEIDCount As Integer = m_FiPEx__1.pPropset.GetProperty("numFlagEIDs")
                If iFlagEIDCount > 0 Then

                    ' Get the active network
                    ' The network shouldn't be empty because the extension wouldn't be enabled otherwise


                    Dim pGeometricNetwork As IGeometricNetwork
                    pGeometricNetwork = m_pNetworkAnalysisExt.CurrentNetwork

                    ' Get the network elements for EID retrieval
                    Dim pNetwork As INetwork
                    Dim pNetElements As INetElements
                    pNetwork = pGeometricNetwork.Network
                    pNetElements = CType(pNetwork, INetElements)

                    Dim pNetworkAnalysisExtFlags As INetworkAnalysisExtFlags
                    pNetworkAnalysisExtFlags = CType(m_UNAExt, INetworkAnalysisExtFlags)

                    ' if current network has no flags loaded (double ensure this isn't run twice - 
                    '  only at the document load)
                    If pNetworkAnalysisExtFlags.JunctionFlagCount = 0 Then

                        Dim i As Integer = 0
                        Dim iEID, iFID, iFCID, iSubID As Integer
                        Dim lFlagEIDs As New List(Of Integer)
                        For i = 0 To iFlagEIDCount - 1
                            iEID = Convert.ToInt32(m_FiPEx__1.pPropset.GetProperty("flagEID" + i.ToString))
                            lFlagEIDs.Add(iEID)
                        Next

                        Dim pFlagDisplay As IFlagDisplay
                        Dim pRgbColor As IRgbColor
                        pRgbColor = New RgbColor

                        ' Set the flag symbol color and parameters
                        Dim pSimpleMarkerSymbol As ISimpleMarkerSymbol
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
                        'm_FlagSymbol4 = pSimpleMarkerSymbol
                        Dim pSymbol As ISymbol
                        Dim pJuncFlagDisplay As IJunctionFlagDisplay

                        i = 0
                        For i = 0 To lFlagEIDs.Count - 1
                            iEID = lFlagEIDs(i)
                            ' just a safe guard against a screwup - ieid = 0
                            If iEID <> 0 Then
                                pNetElements.QueryIDs(iEID, esriElementType.esriETJunction, iFCID, iFID, iSubID)

                                ' Display the flags as a JunctionFlagDisplay type
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
                                pNetworkAnalysisExtFlags.AddJunctionFlag(pJuncFlagDisplay)
                            End If ' ieid = 0
                        Next
                        ' flag count > 0 
                        Dim pActiveView As IActiveView = CType(My.ArcMap.Document.FocusMap, IActiveView)
                        pActiveView.Refresh()
                    End If 'there were any flags set for this network

                End If
                m_FiPEx__1.m_bFlagsLoaded = True
                m_FiPEx__1.pPropset.SetProperty("flagEIDsLoadedyn", m_FiPEx__1.m_bFlagsLoaded)
            End If
        End If

        ' Load the flags saved with the doc if there is a network loaded
        ' and there is a network loaded.  This occurs

        ' Disable Tool if extension is not enabled or
        ' Disable Tool if no networks are loaded

        'If My.ArcMap.Application IsNot Nothing Then
        '    'If m_DFOExt4 = ESRI.ArcGIS.esriSystem.esriExtensionState.esriESDisabled Then
        '    '    Enabled = False
        '    '    Exit Sub
        '    'End If
        '    Dim pUNAExtension As IUtilityNetworkAnalysisExt
        '    pUNAExtension = FishPassageExtension.GetUNAExt

        '    Dim pNetworkAnalysisExt As INetworkAnalysisExt
        '    pNetworkAnalysisExt = CType(pUNAExtension, INetworkAnalysisExt)

        '    ' If there are no networks disable the tool
        '    If pNetworkAnalysisExt.NetworkCount = 0 Then
        '        Enabled = False
        '        Exit Sub
        '    Else
        '        Enabled = True
        '    End If
        'Else
        '    Enabled = False
        'End If

    End Sub
    'Public Overrides Sub OnCreate(ByVal hook As Object)
    '    If Not hook Is Nothing Then
    '        m_application4 = CType(hook, IApplication)

    '        'Disable if it is not ArcMap
    '        If TypeOf hook Is IMxApplication Then
    '            MyBase.m_enabled = True
    '        Else
    '            MyBase.m_enabled = False
    '        End If
    '    End If

    '    ' Obtain a reference to Utility Network Analysis Ext in the current Doc
    '    Dim pUID As New ESRI.ArcGIS.esriSystem.UID
    '    pUID.Value = "{98528F9B-B971-11D2-BABD-00C04FA33C20}"

    '    Dim pExtension As IExtension = m_application4.FindExtensionByCLSID(pUID)
    '    m_UtilityNetworkAnalysisExt4 = CType(pExtension, IUtilityNetworkAnalysisExt)

    '    ' Obtain a reference to the DFOBarriersAnalysis Extension
    '    Dim pUID2 As New ESRI.ArcGIS.esriSystem.UID
    '    pUID2.Value = "FiPEx.DFOBarriersAnalysisExtension"

    '    pExtension = m_application4.FindExtensionByCLSID(pUID2)
    '    m_DFOExt4 = CType(pExtension, FiPEx.DFOBarriersAnalysisExtension)

    '    ' TODO:  Add other initialization code
    'End Sub

    'Public Overrides Sub OnClick()
    '    'TODO: Add PlaceRemoveFlag.OnClick implementation
    'End Sub

    ' Protected Overloads Sub OnMouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Integer, ByVal Y As Integer)
    Protected Overrides Sub OnMouseDown(ByVal arg As MouseEventArgs)

        ' Created By: Greig Oldford
        ' Date: July 4, 2009
        ' Purpose: To place or remove a junction flag and label it
        ' Process Logic:
        ' Get Coordinates
        ' If a junction is within 100 map units
        '   Get the ID of the junction
        '   If the junction already has a flag on it
        '     Remove the flag
        '   Else if the junction does not have a flag on it
        '     add the flag to the network
        '     place a flag symbol on the map
        '     label the flag on the map
        '
        ' NOTE: This tool might be used to place a junction OR an edge flag... maybe?

        ' Get reference to the current network through Utility Network interface

        Dim FiPEx__1 As FishPassageExtension = FishPassageExtension.GetExtension
        Dim pUNAExtension As IUtilityNetworkAnalysisExt
        pUNAExtension = FishPassageExtension.GetUNAExt

        Dim pNetworkAnalysisExt As INetworkAnalysisExt = CType(pUNAExtension, INetworkAnalysisExt)
        Dim pGeometricNetwork As IGeometricNetwork = pNetworkAnalysisExt.CurrentNetwork
        Dim pMxDocument As IMxDocument = CType(My.ArcMap.Application.Document, IMxDocument)
        Dim pMap As IMap = pMxDocument.FocusMap
        Dim pMxApp As IMxApplication = CType(My.ArcMap.Application, IMxApplication)
        Dim pActiveView As IActiveView = CType(pMap, IActiveView)
        Dim pNetwork As INetwork = pGeometricNetwork.Network
        Dim pNetElements As INetElements = CType(pNetwork, INetElements)
        Dim pTraceTasks As ITraceTasks = CType(pUNAExtension, ITraceTasks)
        Dim pNetworkAnalysisExtResults As INetworkAnalysisExtResults = CType(pUNAExtension, INetworkAnalysisExtResults)

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
        ' Get the current network flags
        Dim pNetworkAnalysisExtFlags As INetworkAnalysisExtFlags = CType(pUNAExtension, INetworkAnalysisExtFlags)

        Dim pFlagDisplay As IFlagDisplay
        Dim pOriginaljuncFlagsListGEN As IEnumNetEIDBuilderGEN = New EnumNetEIDArray
        Dim iEID As Integer
        Dim bFlagCheck As Boolean = False
        Dim i As Integer = 0
        ' If there ARE flags
        If pNetworkAnalysisExtFlags.JunctionFlagCount <> 0 Then
            ' Get the EIDs
            For i = 0 To pNetworkAnalysisExtFlags.JunctionFlagCount - 1
                ' Use the bFlagDisplay to retrieve the EIDs of the junction flags
                pFlagDisplay = CType(pNetworkAnalysisExtFlags.JunctionFlag(i), IFlagDisplay)
                iEID = pNetElements.GetEID(pFlagDisplay.FeatureClassID, pFlagDisplay.FID, pFlagDisplay.SubID, esriElementType.esriETJunction)
                If iEID = xEID Then
                    bFlagCheck = True
                End If
                pOriginaljuncFlagsListGEN.Add(iEID)
            Next
        End If

        ' QI to and get an array interface that has 'count' and 'next' methods
        Dim pOriginaljuncFlagsList As IEnumNetEID = CType(pOriginaljuncFlagsListGEN, IEnumNetEID)

        ' ------------------------------------
        ' Add or Remove a flag

        ' Set up the flag marker symbol object
        Dim pSimpleMarkerSymbol As ISimpleMarkerSymbol
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

        Dim pJunctionFlagDisplay As IJunctionFlagDisplay
        Dim netElements As INetElements = CType(pGeometricNetwork.Network, INetElements)
        Dim FCID As Integer, FID As Integer, subID As Integer

        ' convert the EID to a feature class ID, feature ID, and sub ID            
        netElements.QueryIDs(xEID, esriElementType.esriETJunction, FCID, FID, subID)

        ' if there is no flag on the junction add one
        If bFlagCheck = False Then


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
            pNetworkAnalysisExtFlags.AddJunctionFlag(pJunctionFlagDisplay)

            ' Label the flag if needed
            Call LabelFlag(FCID, FID)

            ' If there is already a flag present then remove it
        ElseIf bFlagCheck = True Then

            ' Label the flag if needed
            Call UnLabelFlag(FCID, FID)

            ' Cannot just remove a junction flag
            ' -needs to be two-step process
            pNetworkAnalysisExtFlags.ClearFlags()

            ' Filter the EIDs and make a new list without the clicked junction
            Dim pNewJuncFlagsListGEN As IEnumNetEIDBuilderGEN = New EnumNetEIDArray
            pOriginaljuncFlagsList.Reset()
            For i = 0 To pOriginaljuncFlagsList.Count - 1
                iEID = pOriginaljuncFlagsList.Next
                If iEID <> xEID Then
                    pNewJuncFlagsListGEN.Add(iEID)
                End If
            Next
            Dim pNewJuncFlagsList As IEnumNetEID = CType(pNewJuncFlagsListGEN, IEnumNetEID)

            pNewJuncFlagsList.Reset()
            For i = 0 To pNewJuncFlagsList.Count - 1

                iEID = pNewJuncFlagsList.Next
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
                pNetworkAnalysisExtFlags.AddJunctionFlag(pJunctionFlagDisplay)
            Next

        End If

        ' refresh the view
        pActiveView.Refresh()

    End Sub
    Private Sub LabelFlag(ByVal iFCID As Integer, ByVal iFID As Integer)

        ' Created By: Greig Oldford
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

        Dim FiPEx__1 As FishPassageExtension = FishPassageExtension.GetExtension
        Dim pUNAExtension As IUtilityNetworkAnalysisExt
        pUNAExtension = FishPassageExtension.GetUNAExt


        Dim pMxDocument As IMxDocument = CType(My.ArcMap.Application.Document, IMxDocument)
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
        If FiPEx__1.m_bLoaded = True Then
            ' Get the barrier ID Fields
            iBarrierIDs = Convert.ToInt32(FiPEx__1.pPropset.GetProperty("numBarrierIDs"))
            If iBarrierIDs > 0 Then
                For j = 0 To iBarrierIDs - 1
                    sBarrierIDLayer = Convert.ToString(FiPEx__1.pPropset.GetProperty("BarrierIDLayer" + j.ToString))
                    sBarrierIDField = Convert.ToString(FiPEx__1.pPropset.GetProperty("BarrierIDField" + j.ToString))
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
        'Dim pAnnoLayerProps2 As IAnnotateLayerProperties
        Dim aLELayerProps As ILabelEngineLayerProperties
        Dim pNewLELayerProps As ILabelEngineLayerProperties
        Dim pEnumVisibleElements As IElementCollection
        ' See here for a couple notes: 
        ' http://edndoc.esri.com/arcobjects/9.2/ComponentHelp/esriCarto/IAnnotateLayerProperties.htm
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
        Dim sSearchString As String
        Dim bLabelMatch As Boolean = False
        Dim bUserLabel As Boolean = False
        Dim bString As Boolean = False
        Dim bClassMatch As Boolean = False
        Dim iClassNum As Integer
        Dim pFieldType As esriFieldType
        Dim iFieldIndex As Integer
        Dim pFields As IFields

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
                                                MsgBox("Could not convert label to string for flag. Please check label field. " + _
                                                       " Labelfield: " + sLabelField + ". Label value: " + sLabelValue + ". " + ex.Message)
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

                            'populate visible elements collection
                            'pAnnotateMap = pMap.AnnotationEngine
                            'pAnnotateMapProps = New AnnotateMapProperties
                            'pAnnotateMapProps.AnnotateLayerPropertiesCollection = pAnnoLayerPropsColl
                            'pAnnoLayerProps.FeatureLayer = pGFLayer
                            'pAnnotateMap.Label(pAnnotateMapProps, pMap, Nothing)

                            ' Query the element collection for the visible labels
                            ' And see if there is already a label that matches the 
                            ' one that would be placed.
                            ' If a label field has been found from the extension settings
                            ' then search for that in the visible labels.  If not search 
                            ' for the feature objectID. 
                            'For j = 0 To pEnumVisibleElements.Count - 1

                            '    pEnumVisibleElements.QueryItem(j, pElement)
                            '    pTextElement = CType(pElement, ITextElement)

                            '    If (Not pTextElement Is Nothing) Then
                            '        sText = pTextElement.Text
                            '        If bUserLabel = True Then
                            '            If sLabelValue = sText Then
                            '                bLabelMatch = True
                            '            End If
                            '        Else
                            '            If iFID.ToString = sText Then
                            '                bLabelMatch = True
                            '            End If
                            '        End If
                            '    End If
                            'Next j
                        Next ' PropsIndex

                        ' TEMP SOLUTION
                        ' In where clause from the class, if found, search for the current feature
                        '   Check if using user settings for label to create search string
                        If bClassMatch = True Then
                            If sPreviousWhereClause <> "" Then
                                If bUserLabel = True Then
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
                                MsgBox("An error was encountered trying to label this flag" + ex.Message)
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
    Private Sub UnLabelFlag(ByVal iFCID As Integer, ByVal iFID As Integer)

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

        Dim FiPEx__1 As FishPassageExtension = FishPassageExtension.GetExtension

        Dim pMxDocument As IMxDocument = CType(My.ArcMap.Application.Document, IMxDocument)
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
        If FiPEx__1.m_bLoaded = True Then
            ' Get the barrier ID Fields
            iBarrierIDs = Convert.ToInt32(FiPEx__1.pPropset.GetProperty("numBarrierIDs"))
            If iBarrierIDs > 0 Then
                For j = 0 To iBarrierIDs - 1
                    sBarrierIDLayer = Convert.ToString(FiPEx__1.pPropset.GetProperty("BarrierIDLayer" + j.ToString))
                    sBarrierIDField = Convert.ToString(FiPEx__1.pPropset.GetProperty("BarrierIDField" + j.ToString))
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

                            '' If the layer had no labels already then make sure the new 
                            '' class (default) is defaulted to 'off' - don't show labels
                            '' otherwise it will show labels for the 'default' class. 
                            'If propsIndex = 0 And pAnnoLayerPropsColl.Count = 1 Then
                            '    If pGFLayer.DisplayAnnotation = False Then
                            '        pAnnoLayerProps.DisplayAnnotation = False
                            '    End If
                            'End If

                            'populate visible elements collection
                            'pAnnotateMap = pMap.AnnotationEngine
                            'pAnnotateMapProps = New AnnotateMapProperties
                            'pAnnotateMapProps.AnnotateLayerPropertiesCollection = pAnnoLayerPropsColl
                            'pAnnoLayerProps.FeatureLayer = pGFLayer
                            'pAnnotateMap.Label(pAnnotateMapProps, pMap, Nothing)

                            ' Query the element collection for the visible labels
                            ' And see if there is already a label that matches the 
                            ' one that would be placed.
                            ' If a label field has been found from the extension settings
                            ' then search for that in the visible labels.  If not search 
                            ' for the feature objectID. 
                            'For j = 0 To pEnumVisibleElements.Count - 1

                            '    pEnumVisibleElements.QueryItem(j, pElement)
                            '    pTextElement = CType(pElement, ITextElement)

                            '    If (Not pTextElement Is Nothing) Then
                            '        sText = pTextElement.Text
                            '        If bUserLabel = True Then
                            '            If sLabelValue = sText Then
                            '                bLabelMatch = True
                            '            End If
                            '        Else
                            '            If iFID.ToString = sText Then
                            '                bLabelMatch = True
                            '            End If
                            '        End If
                            '    End If
                            'Next j
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
                        iStringPosit = InStr(sPreviousWhereClause, sSearchString)
                        If iStringPosit <> 0 Then
                            bLabelMatch = True

                            ' need to re-search to check if this clause is found in the middle of the 
                            ' where filter string or on the end (if middle then there is an "OR" to remove)
                            ' Three situations: 1) searchstring is the only one present (no "or")
                            '                   2) searchstring is at the end (preceded by "or")
                            '                   3) searchstring is at the beginning (followed by "or")
                            sORSearchString = "OR " & sSearchString
                            sSearchStringOR = sSearchString & " OR"

                            If InStr(sPreviousWhereClause, sSearchStringOR) <> 0 Then

                                sSearchString = sORSearchString
                                sSQL = sPreviousWhereClause.Remove(iStringPosit - 1, sSearchString.Length + 1)
                                pAnnoLayerPropsColl.QueryItem(iClassNum, pAnnoLayerProps, Nothing, Nothing)
                                pAnnoLayerProps.WhereClause = sSQL

                            ElseIf InStr(sPreviousWhereClause, sORSearchString) <> 0 Then

                                sSearchString = "OR " + sSearchString
                                sSQL = sPreviousWhereClause.Remove(iStringPosit - 5, sSearchString.Length + 1)
                                pAnnoLayerPropsColl.QueryItem(iClassNum, pAnnoLayerProps, Nothing, Nothing)
                                pAnnoLayerProps.WhereClause = sSQL

                                ' otherwise remove the preceding "OR" from this
                            Else
                                sSQL = ""
                                pAnnoLayerPropsColl.QueryItem(iClassNum, pAnnoLayerProps, Nothing, Nothing)
                                pAnnoLayerProps.WhereClause = sSQL
                                pAnnoLayerProps.DisplayAnnotation = False

                            End If

                        End If

                        '' If there was no match between the visible labels and the feature's
                        '' value in the label field
                        'If bLabelMatch = False Then

                        '    ' If we're using a label field specified by the user 
                        '    ' then use that field to label, otherwise use the FID field
                        '    If bUserLabel = True Then
                        '        If bString = True Then
                        '            sSQL = sPreviousWhereClause & sLabelField & " = " & "'" & sLabelValue & "'"
                        '        Else
                        '            sSQL = sPreviousWhereClause & sLabelField & " = " & sLabelValue
                        '        End If
                        '    Else
                        '        strOIDName = pGFLayer.FeatureClass.OIDFieldName
                        '        sSQL = sPreviousWhereClause & strOIDName & " = " & iFID
                        '        sLabelField = strOIDName
                        '    End If

                        '    ' If there was a class match found
                        '    '   Then get that annotation layer properties set
                        '    '    set the 'where' clause
                        '    ' If there was no match found
                        '    '    need to add a class so name it
                        '    '    set the 'where' clause
                        '    ' See here for explanation of new layerprops class
                        '    ' http://resources.esri.com/help/9.3/arcgisengine/dotnet/d3f93845-fedc-42f1-827b-912038c6271b.htm
                        '    If bClassMatch = True Then
                        '        pAnnoLayerPropsColl.QueryItem(iClassNum, pAnnoLayerProps, Nothing, Nothing)
                        '        pAnnoLayerProps.WhereClause = sSQL
                        '    Else
                        '        pNewLELayerProps = New LabelEngineLayerPropertiesClass()
                        '        pAnnoLayerProps = CType(pNewLELayerProps, IAnnotateLayerProperties)
                        '        pAnnoLayerProps.Class = "BarrierOrFlagID"
                        '        pAnnoLayerProps.WhereClause = sSQL
                        '        pAnnoLayerPropsColl.Add(pAnnoLayerProps)
                        '    End If

                        'aLELayerProps = CType(pAnnoLayerProps, ILabelEngineLayerProperties)
                        'aLELayerProps.Expression = "[" & sLabelField & "]"

                        '' Change symbol - add halo
                        'Dim pBOLP As IBasicOverposterLayerProperties = New BasicOverposterLayerProperties
                        'pBOLP.PointPlacementOnTop = False
                        'pBOLP.PointPlacementMethod = esriOverposterPointPlacementMethod.esriAroundPoint
                        'pBOLP.LabelWeight = esriBasicOverposterWeight.esriHighWeight

                        'Dim pTxtSym As ITextSymbol = New ESRI.ArcGIS.Display.TextSymbol
                        'pTxtSym = aLELayerProps.Symbol

                        ''mask color
                        'Dim pRgbColor As IRgbColor = New RgbColor
                        '' Set the barrier symbol color and parameters
                        'With pRgbColor
                        '    .Red = 255
                        '    .Green = 255
                        '    .Blue = 255
                        'End With

                        'Dim pFillSymbol As IFillSymbol = New SimpleFillSymbol
                        'pFillSymbol.Color = pRgbColor

                        'Dim pLineSymbol As ILineSymbol = New SimpleLineSymbol
                        'pLineSymbol.Color = pRgbColor
                        'pFillSymbol.Outline = pLineSymbol

                        'Dim pMask As IMask
                        'pMask = CType(pTxtSym, IMask)

                        'pMask.MaskStyle = esriMaskStyle.esriMSHalo
                        'pMask.MaskSize = 1.8
                        'pMask.MaskSymbol = pFillSymbol

                        'pTxtSym = CType(pMask, ITextSymbol)

                        'aLELayerProps.Symbol = pTxtSym
                        'aLELayerProps.BasicOverposterLayerProperties = pBOLP

                        pGFLayer.DisplayAnnotation = True
                    End If
                End If
            End If

        Next
    End Sub
    Protected Overrides Sub OnMouseMove(ByVal arg As MouseEventArgs)
        'TODO: Add PlaceRemoveFlag.OnMouseMove implementation
    End Sub
    Protected Overrides Sub OnMouseUp(ByVal arg As MouseEventArgs)
        'TODO: Add PlaceRemoveFlag.OnMouseUp implementation
    End Sub

End Class