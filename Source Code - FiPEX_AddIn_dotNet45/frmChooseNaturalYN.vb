Imports System.Runtime.InteropServices
Imports System.Windows.Forms
Imports ESRI.ArcGIS.ArcMapUI
Imports ESRI.ArcGIS.Framework
Imports ESRI.ArcGIS.Carto
Imports ESRI.ArcGIS.Geodatabase

Public Class frmChooseNaturalYN
    Private m_app As IApplication
    Private m_MxDoc As IMxDocument
    Private m_sLayer4 As String ' For layer name
    Public m_BarrierIDObj4 As BarrierIDObj
    Private m_FiPEx As FishPassageExtension

    Public Function Form_Initialize(ByVal m_application As ESRI.ArcGIS.Framework.IApplication) As Boolean
        Try
            'm_app = m_application
            'm_MxDoc = CType(m_app.Document, ESRI.ArcGIS.ArcMapUI.IMxDocument)

            '' Obtain a reference to the DFOBarriersAnalysis Extension
            'Dim pUID2 As New ESRI.ArcGIS.esriSystem.UID
            'pUID2.Value = "FiPEx.DFOBarriersAnalysisExtension"

            'Dim pExtension As ESRI.ArcGIS.esriSystem.IExtension
            'pExtension = m_application.FindExtensionByCLSID(pUID2)
            'm_DFOExt8 = CType(pExtension, FiPEx.DFOBarriersAnalysisExtension)

            m_app = m_application
            m_MxDoc = My.ArcMap.Application.Document
            m_FiPEx = FishPassageExtension.GetExtension


            Return True

        Catch ex As Exception
            System.Windows.Forms.MessageBox.Show(ex.ToString, "Initialize")
            Return False
        End Try
    End Function
    Public Sub New(ByVal sLayer As String, ByVal pBarrierIDObj As BarrierIDObj)
        MyBase.New()
        InitializeComponent()

        m_sLayer4 = sLayer
        'm_sLayerType = sLayerType
        m_BarrierIDObj4 = pBarrierIDObj

    End Sub
    Private Sub Options_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        ' Will load the fields corresponding to the layer selected in the previous form
        ' WILL SEARCH TOC FOR LAYER, STOPPING AT FIRST MATCH - may cause issues if
        ' multiple layers have the same name! possible sol'n: use full path + name when 
        ' referring to layers and saving in pPropSet

        Dim pMap As IMap = m_MxDoc.FocusMap
        Dim m As Integer = 0

        ' Find the corresponding layer in the map
        For m = 0 To pMap.LayerCount - 1
            If pMap.Layer(m).Valid = True Then
                If pMap.Layer(m).Name = m_BarrierIDObj4.Layer Then  'THIS MAY CAUSE ISSUES - see top of sub
                    Exit For
                End If
            End If
        Next

        Dim pLayer As ILayer = pMap.Layer(m)
        Dim pFeatureLayer As IFeatureLayer = CType(pLayer, IFeatureLayer)
        Dim pFeatureClass As IFeatureClass = pFeatureLayer.FeatureClass
        Dim pFields As IFields = pFeatureClass.Fields
        Dim pField As IField

        ' For use finding unique values present in field 
        Dim pFeatureCursor As IFeatureCursor
        Dim pDataStats As IDataStatistics
        'Dim pEnumVar As ESRI.ArcGIS.esriSystem.IEnumVariantSimple
        Dim vVar As Object
        Dim pEnumerator As IEnumerator
        Dim bTFCheck As Boolean

        cboBarrierNaturalYNFields.Items.Add("<None>") 'Add a None option to the hab classes

        ' Add the fields to the combo box
        ' But check to see that the field is of 'string' type
        ' with a length less than 3 (could be strictly set to 1)
        ' and contains only 'T' or 'F' values
        For i As Integer = 0 To pFields.FieldCount - 1
            pField = pFields.Field(i)
            If pField.Type = esriFieldType.esriFieldTypeString And pField.Length < 2 Then

                ' Check that there are only T/F values in field.         
                pFeatureCursor = pFeatureClass.Search(Nothing, False)
                pDataStats = New DataStatistics
                Dim pCursor As ICursor = CType(pFeatureCursor, ICursor)

                With pDataStats
                    .Cursor = pCursor
                    .Field = pField.Name
                    pEnumerator = .UniqueValues
                End With

                pEnumerator.Reset()
                pEnumerator.MoveNext()
                vVar = pEnumerator.Current
                bTFCheck = True

                While Not IsNothing(vVar)
                    If Not vVar.ToString = "T" And Not vVar.ToString = "F" Then
                        bTFCheck = False
                    End If
                    pEnumerator.MoveNext()
                    vVar = pEnumerator.Current
                End While
                If bTFCheck = True Then
                    cboBarrierNaturalYNFields.Items.Add(pField.Name)
                End If

            End If
        Next

        ' For each field in the habitat quantity combo check for match and
        ' if so set it selected

        For i As Integer = 0 To cboBarrierNaturalYNFields.Items.Count - 1
            If cboBarrierNaturalYNFields.Items.Item(i).ToString = m_BarrierIDObj4.NaturalYNField Then
                cboBarrierNaturalYNFields.SelectedIndex = i
            End If
        Next
    End Sub
    ' Use this Predicate object to define search terms and return comparison result
    Private Class CompareLayerNamePred
        Private _name As String
        Public Sub New(ByVal name As String)
            _name = name
        End Sub
        Public Function CompareNames(ByVal obj As BarrierIDObj) As Boolean
            Return (_name = obj.Layer)
        End Function
    End Class
    Private Sub cmdCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdCancel.Click
        Me.Close()
    End Sub

    Private Sub cmdSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdSave.Click
        ' Update BarrierIDObjObject
        Dim iIndex As Integer
        Try
            iIndex = cboBarrierNaturalYNFields.SelectedIndex
            If iIndex <> -1 Then
                m_BarrierIDObj4.NaturalYNField = cboBarrierNaturalYNFields.SelectedItem.ToString
            Else
                m_BarrierIDObj4.NaturalYNField = "Not set"
            End If
        Catch ex As Exception
            MsgBox("Error. Please select a natural barrier yes/no field or click 'cancel'. " + ex.Message)
            Exit Sub
        End Try
        m_BarrierIDObj4.NaturalYNField = cboBarrierNaturalYNFields.SelectedItem.ToString
        Me.Close()
    End Sub
End Class