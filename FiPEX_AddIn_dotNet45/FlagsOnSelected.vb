Imports ESRI.ArcGIS.ArcMapUI
Imports ESRI.ArcGIS.Display
Imports ESRI.ArcGIS.EditorExt
Imports ESRI.ArcGIS.NetworkAnalysis
Imports ESRI.ArcGIS.Framework
Imports ESRI.ArcGIS.Geodatabase
Imports ESRI.ArcGIS.Carto

Public Class FlagsOnSelected
    Inherits ESRI.ArcGIS.Desktop.AddIns.Button
    Private m_UNAExt As IUtilityNetworkAnalysisExt
    Private m_pNetworkAnalysisExt As INetworkAnalysisExt
    Private m_FiPEx__1 As FishPassageExtension
    Public Sub New()

    End Sub

    Protected Overrides Sub OnClick()
        ' ================================================================================
        ' Created by: Greig Oldford
        '       Date: May 1, 2009 (date translated from vb6)
        '    Purpose: Set all selected nodes as flags in the network
        '             Label flags with user-set field if necessary.
        ' ================================================================================

        'Change the mouse cursor to hourglass
        Dim pMouseCursor As IMouseCursor
        pMouseCursor = New MouseCursor
        pMouseCursor.SetCursor(2)

        ' Need these variables to loop through selected layers
        Dim pNetworkAnalysisExtFlags As INetworkAnalysisExtFlags
        Dim lJFlagCountBefore As Integer
        'Dim lJFlagCountAfter As Integer
        Dim pNetworkAnalysisExt As INetworkAnalysisExt
        Dim simpleJunctionFCs As IEnumFeatureClass
        Dim pMxDocument As IMxDocument = CType(My.ArcMap.Application.Document, IMxDocument)
        Dim pMap As IMap
        pMap = pMxDocument.FocusMap
        Dim pActiveView As IActiveView
        pActiveView = CType(pMap, IActiveView)
        Dim pFeatureLayer As IFeatureLayer
        Dim pFeatureClass As IFeatureClass

        'get the number of junction flags currently on the Utility Network Analysis extension
        pNetworkAnalysisExtFlags = CType(m_UNAExt, INetworkAnalysisExtFlags)
        lJFlagCountBefore = pNetworkAnalysisExtFlags.JunctionFlagCount

        'get the current network from the Utility Network Analysis extension
        pNetworkAnalysisExt = CType(m_UNAExt, INetworkAnalysisExt)
        Dim pGeometricNetwork As IGeometricNetwork = m_pNetworkAnalysisExt.CurrentNetwork

        If pGeometricNetwork Is Nothing Then
            MsgBox("There is no current geometric network.")
            Exit Sub
        End If

        'get the simple junction feature classes from this geometric network
        simpleJunctionFCs = pGeometricNetwork.ClassesByType(esriFeatureType.esriFTSimpleJunction)

        ' Loop Counter variables
        Dim i As Integer
        Dim lMaxLayerIndex As Integer

        ' Fields for selected features and count of selected features
        Dim pFeatureSelection As IFeatureSelection
        Dim selectionCount As Integer

        ' The Network for the NetElementBarriers Object needs to be an INetwork
        Dim pNetwork As INetwork
        Dim pGFLayer As IGeoFeatureLayer

        'Dim pLayer As IGeoFeatureLayer
        'Dim strOIDName As String

        'pMxDocument = m_application.Document
        'pNetworkAnalysisExt = m_UtilityNetworkAnalysisExt
        'm_GeometricNetwork = pNetworkAnalysisExt.CurrentNetwork
        '' If somehow, there is no network loaded then display a message
        'If m_GeometricNetwork Is Nothing Then
        '    MsgBox("No Currently Selected Geometric Network.  Add one to the map or select one in the Utility Analyst Toolbar")
        'End If

        'pMap = pMxDocument.FocusMap
        'pActiveView = pMap
        'simpleJunctionFCs = m_GeometricNetwork.ClassesByType(ESRI.ArcGIS.Geodatabase.esriFeatureType.esriFTSimpleJunction)
        lMaxLayerIndex = pMap.LayerCount - 1
        pNetwork = pGeometricNetwork.Network

        ' For each feature in the map
        For i = 0 To lMaxLayerIndex
            If pMap.Layer(i).Valid = True Then
                ' If it's a feature layer then
                If TypeOf pMap.Layer(i) Is IFeatureLayer Then
                    pFeatureLayer = CType(pMap.Layer(i), IFeatureLayer)
                    pGFLayer = CType(pMap.Layer(i), IGeoFeatureLayer)

                    ' If it's a junction then set the type as a simple junction
                    If pFeatureLayer.FeatureClass.FeatureType = esriFeatureType.esriFTSimpleJunction Then

                        simpleJunctionFCs.Reset()
                        pFeatureClass = simpleJunctionFCs.Next

                        ' Cycle through these feature classes
                        Do Until pFeatureClass Is Nothing

                            ' Maybe an unneccesary check? -->
                            If pFeatureClass Is pFeatureLayer.FeatureClass Then

                                ' Count number of selected features in layer
                                pFeatureSelection = CType(pFeatureLayer, IFeatureSelection)
                                selectionCount = pFeatureSelection.SelectionSet.Count

                                ' Debug message box code:
                                'MsgBox("This is the selection count: " + selectionCount.ToString())

                                ' If there are selected features in the layer
                                If selectionCount > 0 Then

                                    ' Debug message box code:
                                    'MsgBox("This is the loop number: " + i.ToString())

                                    'strOIDName = pFeatureLayer.FeatureClass.OIDFieldName
                                    ''strOIDName = "BNumber"
                                    LabelFlags(pFeatureSelection, pGFLayer)
                                    pMxDocument.ActiveView.Refresh()

                                    ' Now we can pass the INetwork to the function to set the barriers
                                    SetSelectedFlags(pFeatureLayer, pNetwork)

                                End If

                                Exit Do
                            End If
                            pFeatureClass = simpleJunctionFCs.Next
                        Loop
                    End If
                End If
            End If
        Next i

    End Sub
    Private Sub SetSelectedFlags(ByVal pFeatureLayer As IFeatureLayer, ByVal network As INetwork)
        ' -------------------------------------------------------
        ' Subroutine:   Set Selected Barriers
        ' Author:       Greig Oldford
        ' Date Created: February 15, 2008
        '               Translated from VB6 to VB.NET April 30, 2009
        ' Description:  This function loops through selected feature
        '               layer features  and sets the selected features
        '               as flags in the geometric network /
        '               utility network.
        ' --------------------------------------------------------

        ' Create the Network Element Barriers Object
        'Dim netJuncBarrs As INetElementBarriersGEN = New NetElementBarriers

        ' Try out SelectionSetBarriers since you can add elements one by one
        'Dim pSelectionSetBarriers As ISelectionSetBarriers = New SelectionSetBarriers

        ' enumerate the number of selected features.
        Dim pFeatureSelection As ESRI.ArcGIS.Carto.IFeatureSelection
        Dim pEnumIDs As ESRI.ArcGIS.Geodatabase.IEnumIDs

        ' EnumIDs array starts at zero so set it to one more
        Dim selectedFID As Integer

        ' Get feature class ID of selected feature class
        Dim selectedFCID As Integer

        ' Set variable to hold EIDs
        Dim pEID As ESRI.ArcGIS.Geodatabase.IEnumNetEID
        Dim pNetElements As ESRI.ArcGIS.Geodatabase.INetElements

        ' Get a reference to the current document, map, and activeview
        Dim pMxDocument As IMxDocument = CType(My.ArcMap.Application.Document, IMxDocument)
        Dim pMap As IMap = pMxDocument.FocusMap
        Dim pActiveView As IActiveView = CType(pMap, IActiveView)


        ' Create variables to hold user ids, user class ids, element id's
        ' of network elements
        Dim lFCID As Integer
        Dim lSubID As Integer
        Dim lFID As Integer
        Dim lEID As Integer

        ' Loop Counters
        Dim i As Long
        'Dim p As Long ' Variable for array
        Dim pEIDCount As Integer

        ' If we are using the netJuncBarrs way then we need an array of element IDs
        'Dim lFIDarray() As Integer

        ' New IFlagdisplay field for displaying flags
        Dim pFlagDisplay As ESRI.ArcGIS.NetworkAnalysis.IFlagDisplay

        ' Obtain reference to current map's utility network and create new
        ' Network Barriers object
        Dim pNetworkAnalysisExtFlags As ESRI.ArcGIS.EditorExt.INetworkAnalysisExtFlags

        Dim flagNumber As Integer


        pFeatureSelection = CType(pFeatureLayer, IFeatureSelection)
        pEnumIDs = pFeatureSelection.SelectionSet.IDs
        pEnumIDs.Reset()
        selectedFID = pEnumIDs.Next
        selectedFCID = pFeatureLayer.FeatureClass.FeatureClassID
        pNetElements = CType(network, INetElements)
        pMap = pMxDocument.FocusMap
        pActiveView = CType(pMap, IActiveView)
        pNetworkAnalysisExtFlags = CType(m_UNAExt, INetworkAnalysisExtFlags)
        Dim pGeometricNetwork As IGeometricNetwork = m_pNetworkAnalysisExt.CurrentNetwork

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


        ' Loop through each selected feature
        Do Until selectedFID < 0

            ' Get corresponding EID(s) of Feature Selection ID
            pEID = pNetElements.GetEIDs(selectedFCID, selectedFID, ESRI.ArcGIS.Geodatabase.esriElementType.esriETJunction)

            ' If the selected features are part of the geometric network then tally them
            If Not pEID Is Nothing Then

                pEID.Reset()
                ' Get a count of element IDs
                pEIDCount = pEID.Count

                ' Set the width of the array
                'lFIDarray = New Integer(pEIDCount) {}

                For i = 1 To pEIDCount

                    'p = i - 1

                    ' Get the element ID of this enumeration
                    lEID = pEID.Next

                    ' Query the corresponding user ID's to the element ID
                    pNetElements.QueryIDs(lEID, ESRI.ArcGIS.Geodatabase.esriElementType.esriETJunction, lFCID, lFID, lSubID)

                    'lFIDarray(p) = lFID
                    'pSelectionSetBarriers.Add(lFCID, lFID)

                    ' Display the barriers as a JunctionFlagDisplay type
                    pFlagDisplay = New ESRI.ArcGIS.NetworkAnalysis.JunctionFlagDisplay
                    Dim pSymbol As ESRI.ArcGIS.Display.ISymbol
                    pSymbol = CType(pFlagSymbol, ESRI.ArcGIS.Display.ISymbol)
                    With pFlagDisplay
                        .FeatureClassID = lFCID
                        .FID = lFID
                        .SubID = lSubID ' This was missing from vb6 code for barriersOnSelected but Present for FlagsOnSelected
                        .Geometry = pGeometricNetwork.GeometryForJunctionEID(lEID)
                        .Symbol = pSymbol
                    End With

                    ' Add the flags to the logical network
                    Dim pJunctionFlagDis As ESRI.ArcGIS.NetworkAnalysis.IJunctionFlagDisplay
                    pJunctionFlagDis = CType(pFlagDisplay, ESRI.ArcGIS.NetworkAnalysis.IJunctionFlagDisplay)
                    pNetworkAnalysisExtFlags.AddJunctionFlag(pJunctionFlagDis)

                Next

                ' Debug MsgBox check:
                'Dim lFIDvalue As Integer
                'For Each lFIDvalue In lFIDarray
                '   MsgBox("this is a value of FID and FCID: " + p.ToString() + ", " + lFCID.ToString())
                'Next lFIDvalue
            End If

            'pEIDs = networkElements.GetEIDs(selectedFCID, pEnumIDs, ESRI.ArcGIS.Geodatabase.esriElementType.esriETJunction

            ' Find next selected feature
            selectedFID = pEnumIDs.Next

        Loop

        ' Clear the selection
        pMap.ClearSelection() ' Not present in original vb6 code

        ' Refresh active view to show new barriers
        pActiveView.Refresh() ' Not present in original vb6 code

        ' Return the total number of barriers set

        flagNumber = pNetworkAnalysisExtFlags.JunctionFlagCount
        'MsgBox ("The number of junction barriers now set: " + CStr(barrNumber))


    End Sub
    Private Sub LabelFlags(ByVal pFeatureSelection As ESRI.ArcGIS.Carto.IFeatureSelection, ByVal pGFLayer As ESRI.ArcGIS.Carto.IGeoFeatureLayer)

        ' Created By: Greig Oldford
        ' Purpose: Label barriers or flags if no label is present
        '          using user-set field from extension settings
        '   
        ' 1.0 Read the extension settings
        ' 2.0 label the flag if needed 
        '
        ' Note: This has been adapted to check for a class called
        '       BarrierOrFlag ID in annotation properties. If it
        '       exists the SQL 'where' statement is saved and added to.
        ' 
        ' Bug Note: Since there are issues with finding visible label
        '           elements (see notes further down) the workaround
        '           shows labels as visible if they have been turned on
        '           and then off in ArcMap.  So if you have turned labels
        '           off they might still be in the annotation properties
        '           and return as 'visible.'

        Dim pMxDocument As IMxDocument = CType(My.ArcMap.Application.Document, IMxDocument)
        Dim pMap As IMap = pMxDocument.FocusMap
        Dim pMxApp As IMxApplication = CType(My.ArcMap.Application, IMxApplication)

        ' ------------------------------------
        ' 1.0 Read Extension Barrier label settings
        Dim lBarrierIDs As List(Of BarrierIDObj) = New List(Of BarrierIDObj)
        Dim pBarrierIDObj As New BarrierIDObj(Nothing, Nothing, Nothing, Nothing, Nothing)

        Dim j As Integer
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
                    ' Object for retrieving flag label fields (do not need other object param so declare 'nothing')
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
        Dim aLELayerProps As ILabelEngineLayerProperties
        Dim pEnumVisibleElements As IElementCollection
        'Dim pElement As IElement
        'Dim pTextElement As ITextElement
        'Dim pAnnotateMapProps As IAnnotateMapProperties
        'Dim pAnnotateMap As IAnnotateMap
        Dim pEnumInVisibleElements As IElementCollection

        ' Default the sLabelField to the OIDFieldName
        Dim sLabelField As String = pGFLayer.FeatureClass.OIDFieldName
        Dim sLabelValue As String
        Dim sSQL As String
        'Dim strOIDName As String
        Dim sPreviousWhereClause As String = ""
        'Dim pFLayer As IFeatureLayer
        'Dim sText As String
        Dim bLabelMatch As Boolean = False
        Dim bUserLabel As Boolean = False
        Dim bString As Boolean = False
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
        ' For each layer in the BarrierID list
        '   If it matches this layer
        '     If the field is found in this layer
        '       Then use this field as a label
        '       If it is a string
        '         Set alert variable to true

        If lBarrierIDs IsNot Nothing Then
            If lBarrierIDs.Count <> 0 Then
                For j = 0 To lBarrierIDs.Count - 1
                    If lBarrierIDs.Item(j).Layer = pGFLayer.Name Then

                        pFeatureClass = pGFLayer.FeatureClass
                        pFields = pFeatureClass.Fields

                        If pFields.FindField(lBarrierIDs.Item(j).Field) <> 0 Then

                            sLabelField = lBarrierIDs.Item(j).Field

                            ' If a field was returned then set the alert variable true
                            If sLabelField <> "Not set" Then

                                bUserLabel = True

                                ' Get the field type of the field because if it is a string
                                ' the sql requires quotation wrappers. 
                                iFieldIndex = pFeatureClass.FindField(sLabelField)
                                pFieldType = pFields.Field(iFieldIndex).Type

                                If pFieldType = esriFieldType.esriFieldTypeString Then
                                    bString = True
                                End If
                            Else
                                bString = False ' because the type of field will be objectid
                            End If
                        End If
                    End If
                Next
            End If
        End If

        ' Get the annotation

        pAnnoLayerPropsColl = pGFLayer.AnnotationProperties
        Dim propsIndex As Integer

        Dim pNewLELayerProps As ILabelEngineLayerProperties

        Dim sSearchString As String
        Dim bClassMatch As Boolean = False
        Dim iClassNum As Integer

        Dim pSelSet As ISelectionSet
        Dim pFeatureCursor As IFeatureCursor

        'Dim iField As Integer

        'Dim bFound As Boolean
        Dim pCursor As ICursor

        ' Get the annotation properties
        For propsIndex = 0 To (pAnnoLayerPropsColl.Count - 1)

            pEnumVisibleElements = New ElementCollection
            pEnumInVisibleElements = New ElementCollection
            pAnnoLayerPropsColl.QueryItem(propsIndex, pAnnoLayerProps, pEnumVisibleElements, pEnumInVisibleElements)

            ' If there is already a class called "BarrierOrFlagID" 
            ' Then get the where clause and save it for later
            If pAnnoLayerProps.Class.ToString = "BarrierOrFlagID" Then
                sPreviousWhereClause = pAnnoLayerProps.WhereClause.ToString
                bClassMatch = True
                sSQL = sPreviousWhereClause
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

            ' POPULATE THE VISIBLE ELEMENTS LIST (WORKAROUND)
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

        Next 'Annotation collection

        ' For each feature in the selection set
        '   If the feature's label value is not found in the existing annotation
        '   label properties 
        '     Add the label to an SQL 'where string'

        pSelSet = pFeatureSelection.SelectionSet
        pSelSet.Search(Nothing, False, pCursor)
        pFeatureCursor = CType(pCursor, IFeatureCursor)

        pFeature = pFeatureCursor.NextFeature

        Do While Not pFeature Is Nothing
            bLabelMatch = False ' reset the label match variable
            pFields = pFeature.Fields
            If bUserLabel = True Then
                Try
                    sLabelValue = Convert.ToString(pFeature.Value(pFields.FindField(sLabelField)))
                Catch ex As Exception
                    MsgBox("Could not convert flag label to type 'string'. " + ex.Message + _
                           " . The feature class: " + pFeatureClass.AliasName + " and the feature: " + _
                           pFeature.OID.ToString)
                    Exit Do
                End Try
            Else
                sLabelValue = pFeature.OID.ToString
                sLabelField = pGFLayer.FeatureClass.OIDFieldName
            End If

            If sLabelValue <> "" Then

                ' In where clause from the class, if found, search for the current feature
                '   Check if using user settings for label to create search string
                If bClassMatch = True Then
                    If sPreviousWhereClause <> "" And sSQL = "" Then
                        If bString = True Then
                            sSearchString = sLabelField & " = '" & sLabelValue & "'"
                        Else
                            sSearchString = sLabelField & " = " & sLabelValue
                        End If
                        If InStr(sPreviousWhereClause, sSearchString) <> 0 Then
                            bLabelMatch = True
                        End If
                    ElseIf sSQL <> "" Then
                        ' on further loops (more features selected in class)
                        If InStr(sSQL, sSearchString) <> 0 Then
                            bLabelMatch = True
                        End If
                    Else
                        bLabelMatch = False
                    End If
                End If

                ' If there was no label for the feature found then
                ' add this feature to the SQL 'Where' clause expression
                If bLabelMatch = False Then
                    ' In case the string is not blank, add an 'or' to it
                    If sSQL <> "" Then
                        sSQL = sSQL & " OR "
                    End If

                    If bString = True Then
                        sSQL = sSQL & sLabelField & " = '" & sLabelValue & "'"
                    Else
                        sSQL = sSQL & sLabelField & " = " & sLabelValue
                    End If
                End If ' bLabelMatch is false

            End If ' if not an empty label value
            pFeature = pFeatureCursor.NextFeature
        Loop

        ' If there was a class match found
        '   Then get that annotation layer properties set
        '    set the 'where' clause
        ' If there was no match found
        '    need to add a class so name it
        '    set the 'where' clause
        ' See here for explanation of new layerprops class
        ' http://resources.esri.com/help/9.3/arcgisengine/dotnet/d3f93845-fedc-42f1-827b-912038c6271b.htm
        If bClassMatch = True Then
            pAnnoLayerPropsColl.QueryItem(iClassNum, pAnnoLayerProps, Nothing, Nothing)
            pAnnoLayerProps.WhereClause = sSQL
            pAnnoLayerProps.DisplayAnnotation = True
        Else
            pNewLELayerProps = New LabelEngineLayerPropertiesClass()
            pAnnoLayerProps = CType(pNewLELayerProps, IAnnotateLayerProperties)
            pAnnoLayerProps.Class = "BarrierOrFlagID"
            pAnnoLayerProps.WhereClause = sSQL
            pAnnoLayerPropsColl.Add(pAnnoLayerProps)
            pAnnoLayerProps.DisplayAnnotation = True
        End If

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
                simpleJunctionFCs = pGeometricNetwork.ClassesByType(ESRI.ArcGIS.Geodatabase.esriFeatureType.esriFTSimpleJunction)
                For i = 0 To pMap.LayerCount - 1
                    If pMap.Layer(i).Valid = True Then
                        ' If it's a feature layer then
                        If TypeOf pMap.Layer(i) Is IFeatureLayer Then
                            pFeatureLayer = CType(pMap.Layer(i), IFeatureLayer)
                            ' If it's a junction then set the type as a simple junction
                            If pFeatureLayer.FeatureClass.FeatureType = esriFeatureType.esriFTSimpleJunction Then
                                simpleJunctionFCs.Reset()
                                pFeatureClass = simpleJunctionFCs.Next
                                ' Cycle through these feature classes
                                Do Until pFeatureClass Is Nothing
                                    ' If there is a match between the map layer FC and the junctionFC
                                    If pFeatureClass Is pFeatureLayer.FeatureClass Then
                                        ' Count number of selected features in layer
                                        pFeatureSelection = CType(pFeatureLayer, IFeatureSelection)
                                        iSelectionCount = pFeatureSelection.SelectionSet.Count
                                        If iSelectionCount > 0 Then
                                            Me.Enabled = True
                                        End If
                                    End If
                                    pFeatureClass = simpleJunctionFCs.Next
                                Loop
                            End If
                        End If
                    End If
                Next
            Else
                Me.Enabled = False
            End If
        Else
            Me.Enabled = False
        End If
        'Me.Enabled = FiPEx__1.HasNetwork

    End Sub
End Class
