Imports System.Runtime.InteropServices
Imports System.Windows.Forms
Imports ESRI.ArcGIS.ArcMapUI
Imports ESRI.ArcGIS.Framework
Imports ESRI.ArcGIS.Carto
Imports ESRI.ArcGIS.Geodatabase

'<ComClass(ChooseHabParam.ClassId, ChooseHabParam.InterfaceId, ChooseHabParam.EventsId), _
' ProgId("FiPEx.ChooseHabParam")> _
Public Class ChooseHabParam

    Private m_app As IApplication
    Private m_MxDoc As IMxDocument
    Private m_sLayer As String ' For layer name
    Private m_sLayerType As String ' For layer type
    Public m_LayerToAdd As LayerToAdd
    Private m_FIPEx As FishPassageExtension
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
            'm_app = m_application
            'm_MxDoc = CType(m_app.Document, ESRI.ArcGIS.ArcMapUI.IMxDocument)

            '' Obtain a reference to the DFOBarriersAnalysis Extension
            'Dim pUID2 As New ESRI.ArcGIS.esriSystem.UID
            'pUID2.Value = "FiPEx.DFOBarriersAnalysisExtension"

            'Dim pExtension As ESRI.ArcGIS.esriSystem.IExtension
            'pExtension = m_application.FindExtensionByCLSID(pUID2)
            'm_DFOExt = CType(pExtension, FiPEx.DFOBarriersAnalysisExtension)

            m_app = m_application
            m_MxDoc = My.ArcMap.Application.Document
            m_FiPEx = FishPassageExtension.GetExtension

            Return True

        Catch ex As Exception
            System.Windows.Forms.MessageBox.Show(ex.ToString, "Initialize")
            Return False
        End Try
    End Function
    Public Sub New(ByVal sLayer As String, ByVal pLayerToAdd As LayerToAdd)
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
            If cmboHabQuanField.Items.Item(i).ToString = m_LayerToAdd.QuanField Then
                cmboHabQuanField.SelectedIndex = i
            End If
        Next
        For i As Integer = 0 To cmboHabClsField.Items.Count - 1
            If cmboHabClsField.Items.Item(i).ToString = m_LayerToAdd.ClsField Then
                cmboHabClsField.SelectedIndex = i
            End If
        Next

        cboUnits.Items.Add("Metres")
        cboUnits.Items.Add("Kilometres")
        cboUnits.Items.Add("Square Metres")
        cboUnits.Items.Add("Hectares")
        cboUnits.Items.Add("Feet")
        cboUnits.Items.Add("Miles")
        cboUnits.Items.Add("Square Miles")
        cboUnits.Items.Add("Acres")
        cboUnits.Items.Add("Hectometres")
        cboUnits.Items.Add("Dekametres")

        For i As Integer = 0 To cboUnits.Items.Count - 1
            If cboUnits.Items.Item(i).ToString = m_LayerToAdd.UnitField Then
                cboUnits.SelectedIndex = i
            End If
        Next

        'End If
        'End If

        'If m_DFOExt.bLoaded = True Then
        '    Dim sPropHabLayerNum As String
        '    Dim sPropHabLayer As String
        '    Dim sPropClsField As String
        '    Dim sPropQuanField As String

        '    If m_sLayerType = "line" Then
        '        sPropHabLayerNum = "numLines"
        '        sPropHabLayer = "incLine"
        '        sPropClsField = "LineClassField"
        '        sPropQuanField = "LineQuanField"
        '    ElseIf m_sLayerType = "poly" Then
        '        sPropHabLayerNum = "numPolys"
        '        sPropHabLayer = "incPoly"
        '        sPropClsField = "PolyClassField"
        '        sPropQuanField = "PolyQuanField"
        '    End If

        '    ' Make sure the current fields set in PropSet are selected
        '    If m_DFOExt.pPropset.GetProperty(sPropHabLayerNum) > 0 Then
        '        For i As Integer = 0 To m_DFOExt.pPropset.GetProperty(sPropHabLayerNum) - 1
        '            If m_DFOExt.pPropset.GetProperty(sPropHabLayer + i) = m_sLayer Then  'if the layer in Propset is the same
        '                With cmboHabClsField
        '                    For j As Integer = 0 To .Items.Count - 1  ' For each field in the combo box
        '                        If .Items.Item(j) = m_DFOExt.pPropset.GetProperty(sPropClsField) Then  'If the field matches the field in propset
        '                            .SelectedIndex = j
        '                        End If
        '                    Next
        '                End With
        '                With cmboHabQuanField
        '                    For k As Integer = 0 To .Items.Count - 1  ' For each field in the combo box
        '                        If .Items.Item(k) = m_DFOExt.pPropset.GetProperty(sPropClsField) Then  'If the field matches the field in propset
        '                            .SelectedIndex = k
        '                        End If
        '                    Next
        '                End With
        '            End If

        '        Next
        '    End If
        'End If


    End Sub

    Private Sub cmdSaveHabSettings_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdSaveHabSettings.Click
        ' Update LayerToAddObject

        Dim iIndex As Integer
        Try
            iIndex = cmboHabClsField.SelectedIndex
            If iIndex <> -1 Then
                m_LayerToAdd.ClsField = cmboHabClsField.SelectedItem.ToString
            Else
                m_LayerToAdd.ClsField = "<None>"
            End If
        Catch ex As Exception
            MsgBox("Error on save.  Please select a habitat class field.  Select 'none' if no classes. " + ex.Message)
            Exit Sub
        End Try

        Try
            iIndex = cmboHabQuanField.SelectedIndex
            If iIndex <> -1 Then
                m_LayerToAdd.QuanField = cmboHabQuanField.SelectedItem.ToString
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
                m_LayerToAdd.UnitField = cboUnits.SelectedItem.ToString
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
        Public Function CompareNames(ByVal obj As LayerToAdd) As Boolean
            Return (_name = obj.Layer)
        End Function
    End Class
    Private Sub cmdCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdCancel.Click
        Me.Close()
    End Sub
End Class