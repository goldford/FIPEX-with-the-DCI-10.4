Imports ESRI.ArcGIS.Framework
Imports ESRI.ArcGIS.ArcMapUI
Imports ESRI.ArcGIS.Carto
Imports ESRI.ArcGIS.Geodatabase
Imports ESRI.ArcGIS.EditorExt

Public Class frmDistanceToMouth
    Public m_bCancel1 As Boolean
    Public m_bUseFiPExQuan As Boolean
    Public m_lLayersAndFCIDs As New List(Of LayersAndFCIDs)
    Private m_MxDoc As IMxDocument
    Private m_FiPEx As FishPassageExtension
    Private m_UNAExt As IUtilityNetworkAnalysisExt
    Private m_pNetworkAnalysisExt As INetworkAnalysisExt

    Private Sub frmDistanceToMouth_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        ' loop through map document and load layers that 
        ' (a) have a selection
        ' (b) are simple edge or junction features in the active geometric network

        Dim pMap As IMap
        If m_MxDoc Is Nothing Then
            m_MxDoc = My.ArcMap.Application.Document
        End If

        pMap = m_MxDoc.FocusMap

        If m_FiPEx Is Nothing Then
            m_FiPEx = FiPEX_ArcMap_10p4_up_AddIn_dotNet45_2020.FishPassageExtension.GetExtension()
        End If
        If m_UNAExt Is Nothing Then
            m_UNAExt = FiPEX_ArcMap_10p4_up_AddIn_dotNet45_2020.FishPassageExtension.GetUNAExt
        End If
        If m_pNetworkAnalysisExt Is Nothing Then
            m_pNetworkAnalysisExt = CType(m_UNAExt, INetworkAnalysisExt)
        End If

        Dim pLayersAndFCIDs As New LayersAndFCIDs(Nothing, Nothing)

        Dim pGeometricNetwork As IGeometricNetwork = m_pNetworkAnalysisExt.CurrentNetwork
        Dim i, iSelectionCount As Integer
        Dim pFeatureLayer As IFeatureLayer
        Dim pFeatureClass As IFeatureClass
        Dim pFeatureSelection As IFeatureSelection
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
                                    ' check if list contains a layer by this name and the unique FC ID is not in the running list
                                    pLayersAndFCIDs = New LayersAndFCIDs(pMap.Layer(i).Name, pFeatureLayer.FeatureClass.FeatureClassID)
                                    If Not chklistSimpleFCs.Items.Contains(pMap.Layer(i).Name) And Not m_lLayersAndFCIDs.Contains(pLayersAndFCIDs) Then
                                        chklistSimpleFCs.Items.Add(pMap.Layer(i).Name)
                                        m_lLayersAndFCIDs.Add(pLayersAndFCIDs)
                                    End If
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
                                    pLayersAndFCIDs = New LayersAndFCIDs(pMap.Layer(i).Name, pFeatureLayer.FeatureClass.FeatureClassID)
                                    If Not chklistSimpleFCs.Items.Contains(pMap.Layer(i).Name) And Not m_lLayersAndFCIDs.Contains(pLayersAndFCIDs) Then
                                        chklistSimpleFCs.Items.Add(pMap.Layer(i).Name)
                                        m_lLayersAndFCIDs.Add(pLayersAndFCIDs)
                                    End If
                                End If
                            End If
                            pFeatureClass = simpleEdgeFCs.Next
                        Loop
                    End If
                End If
            End If
        Next

    End Sub

    Private Sub cmdCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdCancel.Click
        m_bCancel1 = True
        Me.Close()
    End Sub
    Public Function Form_Initialize(ByVal m_application As ESRI.ArcGIS.Framework.IApplication) As Boolean
        Try

            m_MxDoc = My.ArcMap.Application.Document
            m_FiPEx = FishPassageExtension.GetExtension

            Return True

        Catch ex As Exception
            System.Windows.Forms.MessageBox.Show(ex.ToString, "Initialize")
            Return False
        End Try
    End Function

    Private Sub cmdDoIt_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdDoIt.Click
        m_bUseFiPExQuan = chkUseFiPExQuantity.CheckState

        Dim i As Integer = 0
        Dim k As Integer = 0
        Dim m As Integer = 0
        Dim iFCID As Integer

        Dim pMap As IMap
        If m_MxDoc Is Nothing Then
            m_MxDoc = My.ArcMap.Application.Document
        End If

        ' get the FCIDs of the checked network feature classes
        Dim pFeatureLayer As IFeatureLayer
        Dim pFeatureClass As IFeatureClass
        Dim sLayerName As String
        ' Dim pFeatureSelection As IFeatureSelection
        'Dim simpleJunctionFCs As IEnumFeatureClass
        'Dim simpleEdgeFCs As IEnumFeatureClass
        'Dim pGeometricNetwork As IGeometricNetwork = m_pNetworkAnalysisExt.CurrentNetwork
        'simpleJunctionFCs = pGeometricNetwork.ClassesByType(ESRI.ArcGIS.Geodatabase.esriFeatureType.esriFTSimpleJunction)
        ' simpleEdgeFCs = pGeometricNetwork.ClassesByType(ESRI.ArcGIS.Geodatabase.esriFeatureType.esriFTSimpleEdge)

        pMap = m_MxDoc.FocusMap

        For k = 0 To chklistSimpleFCs.Items.Count - 1
            If Not chklistSimpleFCs.GetItemChecked(k) Then
                sLayerName = chklistSimpleFCs.Items.Item(k)
                ' get index of first occurrence of the layer lame
                m = m_lLayersAndFCIDs.Count - 1
                For m = m_lLayersAndFCIDs.Count - 1 To 0 Step -1
                    Dim bMatch As Boolean = False
                    If m_lLayersAndFCIDs.Item(m).LayerName = sLayerName Then
                        m_lLayersAndFCIDs.RemoveAt(m)
                    End If
                Next
            End If
        Next
        Me.Close()
    End Sub
    Private Class FindFeatureLayerNamePredicate
        ' this class should help return a 
        ' list object 
        Private _layername As String

        Public Sub New(ByVal layername As String)
            Me._layername = layername
        End Sub

        Public Function CompareLayerName(ByVal obj As LayersAndFCIDs) As Boolean
            Return (_layername = obj.LayerName)
        End Function
    End Class


End Class