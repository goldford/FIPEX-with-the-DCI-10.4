Imports ESRI.ArcGIS.EditorExt
Imports ESRI.ArcGIS.NetworkAnalysis
Imports ESRI.ArcGIS.Geodatabase
Imports ESRI.ArcGIS.ArcMapUI
Imports ESRI.ArcGIS.Carto

Public Class ClearFlagsAndLabels
    Inherits ESRI.ArcGIS.Desktop.AddIns.Button
    Private m_UNAExt As IUtilityNetworkAnalysisExt
    Private m_pNetworkAnalysisExt As INetworkAnalysisExt
    Private m_FiPEx__1 As FishPassageExtension

    Public Sub New()

    End Sub

    Protected Overrides Sub OnClick()
        'TODO: Add ClearBarriersAndLabels.OnClick implementation
        ' =============================================================
        '
        ' Created by: Greig Oldford
        ' Date last modified: Oct. 21, 2010
        ' Purpose: Clear the barriers and labels in the current network
        '
        '
        ' =============================================================

        ' (get network)


        ' Get reference to the current network through Utility Network interface
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


        ' 1. Get network barriers


        ' ===============GET CURRENT FLAGS AND BARRIERS ==================
        ' Before all current flags are cleared

        Dim pOriginaljuncFlagsListGEN As IEnumNetEIDBuilderGEN
        Dim pOriginalBarriersListGEN As IEnumNetEIDBuilderGEN
        Dim pOriginalEdgeFlagsListGEN As IEnumNetEIDBuilderGEN

        Dim pNetworkAnalysisExtBarriers As INetworkAnalysisExtBarriers
        Dim pNetworkAnalysisExtFlags As INetworkAnalysisExtFlags
        Dim pFlagDisplay As IFlagDisplay
        Dim bEID, i As Integer

        ' QI the Flags and barriers
        pNetworkAnalysisExtFlags = CType(m_UNAExt, INetworkAnalysisExtFlags)
        pNetworkAnalysisExtBarriers = CType(m_UNAExt, INetworkAnalysisExtBarriers)

        ' Was using ArrayList because of advantage of 'count' and 'add' properties
        ' but EnumNetEIDBuilderGEN addition to 9.2 has this functionality
        pOriginalBarriersListGEN = New EnumNetEIDArray
        pOriginalEdgeFlagsListGEN = New EnumNetEIDArray
        pOriginaljuncFlagsListGEN = New EnumNetEIDArray

        ' Save the barriers
        For i = 0 To pNetworkAnalysisExtBarriers.JunctionBarrierCount - 1
            ' Use bFlagDisplay to retrieve EIDs of the barriers for later
            pFlagDisplay = CType(pNetworkAnalysisExtBarriers.JunctionBarrier(i), IFlagDisplay)
            bEID = pNetElements.GetEID(pFlagDisplay.FeatureClassID, pFlagDisplay.FID, pFlagDisplay.SubID, esriElementType.esriETJunction)
            pOriginalBarriersListGEN.Add(bEID)
            'originalBarriersList(i) = bEID
        Next

        ' No edge flag support yet
        Dim pOriginalBarriersList, pOriginalBarriersListSaved As IEnumNetEID

        ' QI to and get an array object that has 'count' and 'next' methods
        pOriginalBarriersList = CType(pOriginalBarriersListGEN, IEnumNetEID)
        ' Save the flags
        i = 0
        For i = 0 To pNetworkAnalysisExtFlags.JunctionFlagCount - 1
            ' Use the bFlagDisplay to retrieve the EIDs of the junction flags
            pFlagDisplay = CType(pNetworkAnalysisExtFlags.JunctionFlag(i), IFlagDisplay)
            bEID = pNetElements.GetEID(pFlagDisplay.FeatureClassID, pFlagDisplay.FID, pFlagDisplay.SubID, esriElementType.esriETJunction)
            pOriginaljuncFlagsListGEN.Add(bEID)
            'pOriginaljuncFlagsList(i) = bEID
        Next

        Dim pOriginaljuncFlagsList As IEnumNetEID
        ' QI to and get an array interface that has 'count' and 'next' methods
        pOriginaljuncFlagsList = CType(pOriginaljuncFlagsListGEN, IEnumNetEID)

        ' ******** NO EDGE FLAG SUPPORT YET *********
        i = 0
        For i = 0 To pNetworkAnalysisExtFlags.EdgeFlagCount - 1

            ' Use the bFlagDisplay to retrieve EIDs of the Edge flags for later
            pFlagDisplay = CType(pNetworkAnalysisExtFlags.EdgeFlag(i), IFlagDisplay)
            bEID = pNetElements.GetEID(pFlagDisplay.FeatureClassID, pFlagDisplay.FID, pFlagDisplay.SubID, esriElementType.esriETJunction)
            pOriginalEdgeFlagsListGEN.Add(bEID)
            'pOriginalEdgeFlagsList(i) = bEID
        Next

        Dim pOriginalEdgeFlagsList As IEnumNetEID
        ' QI to and get an array interface that has 'count' and 'next' methods
        pOriginalEdgeFlagsList = CType(pOriginalEdgeFlagsListGEN, IEnumNetEID)


        ' 3. Cross ref and see which overlaps
        ' 4. Make list of non-overlappping barriers
        ' 5. Clear those labels

        ' ============== UnLABEL FLAGS ===================
        Dim sFlagCheck As String
        Dim iFlagEID, m, FCID, FID, subID As Integer
        Dim flagBarrier As Boolean

        ' Check if flag is on a barrier
        pOriginaljuncFlagsList.Reset()
        For i = 0 To pOriginaljuncFlagsList.Count - 1

            iFlagEID = pOriginaljuncFlagsList.Next
            flagBarrier = False     ' assume flag is not on barrier
            m = 0

            pOriginalBarriersList.Reset()
            For m = 0 To pOriginalBarriersList.Count - 1
                If iFlagEID = pOriginalBarriersList.Next Then
                    flagBarrier = True
                End If
            Next

            ' unlabel if not over barrier
            If flagBarrier = False Then
                pNetElements.QueryIDs(iFlagEID, esriElementType.esriETJunction, FCID, FID, subID)
                UnLabelFlag(FCID, FID)
            End If
        Next

        ' 6. Clear all barriers 

        ' clear current barriers
        pNetworkAnalysisExtFlags.ClearFlags()


        ' refresh the view
        pActiveView.Refresh()
    End Sub
    Private Sub UnLabelFlag(ByVal iFCID As Integer, ByVal iFID As Integer)

        ' Created By: Greig Oldford
        ' Date: July 5, 2009
        ' Purpose: unLabel flags if label is present and not on a barrier
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
        Dim sSQL As String
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

    Protected Overrides Sub OnUpdate()
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
            Dim pNetworkAnalysisExtFlags As INetworkAnalysisExtFlags = CType(m_UNAExt, INetworkAnalysisExtFlags)
            If pNetworkAnalysisExtFlags.JunctionFlagCount > 0 Then
                Me.Enabled = True
            Else
                Me.Enabled = False
            End If
        Else
            Me.Enabled = False
        End If
    End Sub
End Class
