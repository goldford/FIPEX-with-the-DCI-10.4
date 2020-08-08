
Imports System.Runtime.InteropServices
Imports System.Windows.Forms
Imports ESRI.ArcGIS.ArcMapUI
Imports ESRI.ArcGIS.Framework
Imports ESRI.ArcGIS.Carto
Imports ESRI.ArcGIS.Geodatabase

    '<ComClass(ChooseHabParam.ClassId, ChooseHabParam.InterfaceId, ChooseHabParam.EventsId), _
' ProgId("FiPEx.ChooseHabParam")> _

Public Class frmChooseBarrierPerm


    ' ----------------------------------------------------------------------------------------
    ' 1 Get the selected layer in the listbox as a variable
    ' 2 Clear the listbox containing barrier layer Fields
    ' 3 Populate the listbox with the layer's fields
    ' 4.1 If settings from this extension have been loaded then
    '       4.1.1 If there are any layers in the global variable
    '         4.1.1.1 For each of those layers
    '           4.1.1.1 If the layer in the global variable list matches the selected layer
    '           4.1.1.2 Set a module level variable containing all parameters for that field
    '                (that can be used to be passed to the ChooseHabPAram form)
    '           6.2a Load the class, quantity, and unit fields into the appropriate listboxes
    '           6.3a Check a variable saying there has been a match found
    '           6.4a Exit this for loop to avoid multiple adds (see issue outlined above)
    '  7.0 If nothing has been added to the class, quantity, and unit listboxes
    '  7.1 Set a module level variable containing all parameters for that field
    '     (that can be used to be passed to the ChooseHabPAram form)
    '  7.2 Update the listboxes with 'not set'
    ' -----------------------------------------------------------------------------------------
    Private m_app As IApplication
    Private m_MxDoc As IMxDocument
    Private m_sLayer3 As String ' For layer name
    Public m_BarrierIDObj3 As BarrierIDObj
    Private m_FiPEx As FishPassageExtension

    'Private m_OptionsForm As FiPEx.Options

    '#Region "COM GUIDs"
    '    ' These  GUIDs provide the COM identity for this class 
    '    ' and its COM interfaces. If you change them, existing 
    '    ' clients will no longer be able to access the class.
    '    Public Const ClassId As String = "176bc803-4712-418e-9d48-11003992b98e"
    '    Public Const InterfaceId As String = "4a6169f2-d3a6-4cf5-b578-1dee0b1d97b3"
    '    Public Const EventsId As String = "1c82d668-2cba-49b4-a407-354915ecb000"
    '#End Region

    Public Function Form_Initialize(ByVal m_application As ESRI.ArcGIS.Framework.IApplication) As Boolean
        Try
            m_app = m_application
            m_MxDoc = My.ArcMap.Application.Document
            m_FiPEx = FishPassageExtension.GetExtension
            'm_UtilityNetworkAnalysisExt = FishPassageExtension.GetUNAExt

            ' Obtain a reference to the DFOBarriersAnalysis Extension
            'Dim pUID2 As New ESRI.ArcGIS.esriSystem.UID
            'pUID2.Value = "FiPEx.DFOBarriersAnalysisExtension"

            'Dim pExtension As ESRI.ArcGIS.esriSystem.IExtension
            'pExtension = m_application.FindExtensionByCLSID(pUID2)
            'm_DFOExt7 = CType(pExtension, FiPEx.DFOBarriersAnalysisExtension)

            'm_app = m_application
            'pMxDoc = CType(m_app.Document, ESRI.ArcGIS.ArcMapUI.IMxDocument)
           

            ' Obtain a reference to Utility Network Analysis Ext in the current Doc
            'Dim pUID As New ESRI.ArcGIS.esriSystem.UID
            'pUID.Value = "{98528F9B-B971-11D2-BABD-00C04FA33C20}"

            'Dim pExtension As ESRI.ArcGIS.esriSystem.IExtension = m_application.FindExtensionByCLSID(pUID)
            'm_UtilityNetworkAnalysisExt = CType(pExtension, IUtilityNetworkAnalysisExt)



            Return True

        Catch ex As Exception
            System.Windows.Forms.MessageBox.Show(ex.ToString, "Initialize")
            Return False
        End Try
    End Function
    Public Sub New(ByVal sLayer As String, ByVal pBarrierIDObj As BarrierIDObj)
        MyBase.New()
        InitializeComponent()

        m_sLayer3 = sLayer
        'm_sLayerType = sLayerType
        m_BarrierIDObj3 = pBarrierIDObj

    End Sub
    Private Sub Options_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        ' Will load the fields corresponding to the layer selected in the previous form
        ' WILL SEARCH TOC FOR LAYER, STOPPING AT FIRST MATCH - may cause issues if
        ' multiple layers have the same name! 

        Dim pMap As IMap = m_MxDoc.FocusMap
        Dim m As Integer = 0

        ' Find the corresponding layer in the map
        For m = 0 To pMap.LayerCount - 1
            If pMap.Layer(m).Valid = True Then
                If pMap.Layer(m).Name = m_BarrierIDObj3.Layer Then  'THIS MAY CAUSE ISSUES - see top of sub
                    Exit For
                End If
            End If
        Next

        Dim pLayer As ILayer = pMap.Layer(m)
        Dim pFeatureLayer As IFeatureLayer = CType(pLayer, IFeatureLayer)
        Dim pFeatureClass As IFeatureClass = pFeatureLayer.FeatureClass
        Dim pFields As IFields = pFeatureClass.Fields
        Dim pField As IField

        cboBarrierPermFields.Items.Add("<None>") 'Add a None option to the hab classes

        ' Add the fields to the combo box
        For i As Integer = 0 To pFields.FieldCount - 1
            pField = pFields.Field(i)
            ' Field must be of type double. 
            If pField.Type = esriFieldType.esriFieldTypeDouble Then
                cboBarrierPermFields.Items.Add(pField.Name)
            End If
        Next

        ' For each field in the habitat quantity combo check for match and
        ' if so set it selected

        For i As Integer = 0 To cboBarrierPermFields.Items.Count - 1
            If cboBarrierPermFields.Items.Item(i).ToString = m_BarrierIDObj3.PermField Then
                cboBarrierPermFields.SelectedIndex = i
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
            iIndex = cboBarrierPermFields.SelectedIndex
            If iIndex <> -1 Then
                m_BarrierIDObj3.PermField = cboBarrierPermFields.SelectedItem.ToString
            Else
                m_BarrierIDObj3.PermField = "Not set"
            End If
        Catch ex As Exception
            MsgBox("Error.  Please select a permeability field or click 'cancell'. " + ex.Message)
            Exit Sub
        End Try
        Me.Close()
    End Sub
End Class


