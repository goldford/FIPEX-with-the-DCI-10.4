Imports System.Runtime.InteropServices
Imports System.Windows.Forms
Imports ESRI.ArcGIS.ArcMapUI
Imports ESRI.ArcGIS.Framework
Imports ESRI.ArcGIS.Carto
Imports ESRI.ArcGIS.Geodatabase

Public Class ChoosePolyHabParam

    Private m_app As IApplication
    Private m_MxDoc As IMxDocument
    Private m_sLayer As String ' For layer name
    Private m_sLayerType As String ' For layer type
    Public m_LayerToAdd As PolyLayerToAdd
    Private m_FIPEx As FishPassageExtension
   

    Public Function Form_Initialize(ByVal m_application As ESRI.ArcGIS.Framework.IApplication) As Boolean
        Try

            m_app = m_application
            m_MxDoc = My.ArcMap.Application.Document
            m_FiPEx = FishPassageExtension.GetExtension

            Return True

        Catch ex As Exception
            System.Windows.Forms.MessageBox.Show(ex.ToString, "Initialize")
            Return False
        End Try
    End Function
    Public Sub New(ByVal sLayer As String, ByVal pLayerToAdd As PolyLayerToAdd)
        MyBase.New()
        InitializeComponent()

        m_sLayer = sLayer
        'm_sLayerType = sLayerType
        m_LayerToAdd = pLayerToAdd

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
                If pMap.Layer(m).Name = m_LayerToAdd.Layer Then Exit For 'THIS MAY CAUSE ISSUES - see top of sub
            End If
        Next

        Dim pLayer As ILayer = pMap.Layer(m)
        Dim pFeatureLayer As IFeatureLayer = CType(pLayer, IFeatureLayer)
        Dim pFeatureClass As IFeatureClass = pFeatureLayer.FeatureClass
        Dim pFields As IFields = pFeatureClass.Fields
        Dim pField As IField

        cmboHabClsField.Items.Add("<None>") 'Add a None option to the hab classes

        ' Add the fields to the combo box
        For i As Integer = 0 To pFields.FieldCount - 1
            pField = pFields.Field(i)
            cmboHabClsField.Items.Add(pField.Name)
            cmboHabQuanField.Items.Add(pField.Name)
        Next

        ' Get the property of the options form
        'Dim pLayersFields As List(Of LayerToAdd) = New List(Of LayerToAdd)
        'If m_sLayerType = "line" Then
        '    'pLayersFields = m_OptionsForm.LHabParamList
        'ElseIf m_sLayerType = "poly" Then
        '    'pLayersFields = m_OptionsForm.PHabParamList
        'End If

        ' If there are some things in the list
        'If pLayersFields.Count > 0 Then
        '    ' look for a match in the list with the current layer
        '    Dim comparer As New CompareLayerNamePred(m_sLayer)
        '    Dim LayerMatched As LayerToAdd
        '    LayerMatched = pLayersFields.Find(AddressOf comparer.CompareNames)

        ' If there is a match then set those fields as selected
        'If LayerMatched IsNot Nothing Then
        ' For each field in the habitat quantity combo check for match and
        ' if so set it selected

        For i As Integer = 0 To cmboHabQuanField.Items.Count - 1
            If cmboHabQuanField.Items.Item(i).ToString = m_LayerToAdd.HabQuanField Then
                cmboHabQuanField.SelectedIndex = i
            End If
        Next
        For i As Integer = 0 To cmboHabClsField.Items.Count - 1
            If cmboHabClsField.Items.Item(i).ToString = m_LayerToAdd.HabClsField Then
                cmboHabClsField.SelectedIndex = i
            End If
        Next

        cboUnits.Items.Add("Metres")
        cboUnits.Items.Add("Kilometres")
        cboUnits.Items.Add("Feet")
        cboUnits.Items.Add("Miles")
        cboUnits.Items.Add("Square Miles")
        cboUnits.Items.Add("Acres")
        cboUnits.Items.Add("Square Metres")
        cboUnits.Items.Add("Hectares")
        cboUnits.Items.Add("Square Kilometres")
        cboUnits.Items.Add("Hectometres")
        cboUnits.Items.Add("Dekametres")

        For i As Integer = 0 To cboUnits.Items.Count - 1
            If cboUnits.Items.Item(i).ToString = m_LayerToAdd.HabUnitField Then
                cboUnits.SelectedIndex = i
            End If
        Next


    End Sub

    Private Sub cmdSaveHabSettings_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdSaveHabSettings.Click
        ' Update LayerToAddObject

        Dim iIndex As Integer
        Try
            iIndex = cmboHabClsField.SelectedIndex
            If iIndex <> -1 Then
                m_LayerToAdd.HabClsField = cmboHabClsField.SelectedItem.ToString
            Else
                m_LayerToAdd.HabClsField = "<None>"
            End If
        Catch ex As Exception
            MsgBox("Error on save.  Please select a habitat class field.  Select 'none' if no classes. " + ex.Message)
            Exit Sub
        End Try

        Try
            iIndex = cmboHabQuanField.SelectedIndex
            If iIndex <> -1 Then
                m_LayerToAdd.HabQuanField = cmboHabQuanField.SelectedItem.ToString
            Else
                MsgBox("Please select a quantity field or click 'cancel'.")
                Exit Sub
            End If
        Catch ex As Exception
            MsgBox("Error on save.  Please select a habitat quantity field or hit 'cancel' " + ex.Message)
            Exit Sub
        End Try

        Try
            iIndex = cboUnits.SelectedIndex
            If iIndex <> -1 Then
                m_LayerToAdd.HabUnitField = cboUnits.SelectedItem.ToString
            Else
                MsgBox("Please select a units field or click 'cancel'.")
                Exit Sub
            End If
        Catch ex As Exception
            MsgBox("Error on save.  Please select a quantity unit field or hit 'cancel' " + ex.Message)
            Exit Sub
        End Try


        Me.Close()
    End Sub

    ' Use this Predicate object to define search terms and return comparison result
    Private Class CompareLayerNamePred
        Private _name As String
        Public Sub New(ByVal name As String)
            _name = name
        End Sub
        Public Function CompareNames(ByVal obj As PolyLayerToAdd) As Boolean
            Return (_name = obj.Layer)
        End Function
    End Class

    Private Sub cmdCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdCancel.Click
        Me.Close()
    End Sub
End Class