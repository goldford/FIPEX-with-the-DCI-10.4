Imports System.Runtime.InteropServices
Imports System.Windows.Forms
Imports ESRI.ArcGIS.ArcMapUI
Imports ESRI.ArcGIS.Framework
Imports ESRI.ArcGIS.Carto
Imports ESRI.ArcGIS.Geodatabase


Public Class ChooseLineHabParam
    Private m_app As IApplication
    Private m_MxDoc As IMxDocument
    Private m_sLayer As String ' For layer name
    Private m_sLayerType As String ' For layer type
    Public m_LayerToAdd As LineLayerToAdd
    Private m_FIPEx As FishPassageExtension

    Public Function Form_Initialize(ByVal m_application As ESRI.ArcGIS.Framework.IApplication) As Boolean
        Try
            m_app = m_application
            m_MxDoc = My.ArcMap.Application.Document
            m_FIPEx = FishPassageExtension.GetExtension

            Return True

        Catch ex As Exception
            System.Windows.Forms.MessageBox.Show(ex.ToString, "Initialize")
            Return False
        End Try
    End Function

    Public Sub New(ByVal sLayer As String, ByVal pLayerToAdd As LineLayerToAdd)
        MyBase.New()
        InitializeComponent()

        m_sLayer = sLayer
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

        cboHabClassField.Items.Add("<None>") 'Add a None option to the hab classes

        ' Add the fields to the combo box
        For i As Integer = 0 To pFields.FieldCount - 1
            pField = pFields.Field(i)
            cboLengthField.Items.Add(pField.Name)
            cboHabClassField.Items.Add(pField.Name)
            cmboHabQuanField.Items.Add(pField.Name)
        Next

        For i As Integer = 0 To cboLengthField.Items.Count - 1
            If cboLengthField.Items.Item(i).ToString = m_LayerToAdd.LengthField Then
                cboLengthField.SelectedIndex = i
            End If
        Next

        For i As Integer = 0 To cmboHabQuanField.Items.Count - 1
            If cmboHabQuanField.Items.Item(i).ToString = m_LayerToAdd.HabQuanField Then
                cmboHabQuanField.SelectedIndex = i
            End If
        Next
        For i As Integer = 0 To cboHabClassField.Items.Count - 1
            If cboHabClassField.Items.Item(i).ToString = m_LayerToAdd.HabClsField Then
                cboHabClassField.SelectedIndex = i
            End If
        Next

        cboLengthUnits.Items.Add("Metres")
        cboLengthUnits.Items.Add("Kilometres")
        cboLengthUnits.Items.Add("Feet")
        cboLengthUnits.Items.Add("Miles")
        cboLengthUnits.Items.Add("Hectometres")
        cboLengthUnits.Items.Add("Dekametres")

        For i As Integer = 0 To cboLengthUnits.Items.Count - 1
            If cboLengthUnits.Items.Item(i).ToString = m_LayerToAdd.LengthUnits Then
                cboLengthUnits.SelectedIndex = i
            End If
        Next

        cboHabUnits.Items.Add("Metres")
        cboHabUnits.Items.Add("Kilometres")
        cboHabUnits.Items.Add("Feet")
        cboHabUnits.Items.Add("Miles")
        cboHabUnits.Items.Add("Square Miles")
        cboHabUnits.Items.Add("Acres")
        cboHabUnits.Items.Add("Square Metres")
        cboHabUnits.Items.Add("Hectares")
        cboHabUnits.Items.Add("Square Kilometres")
        cboHabUnits.Items.Add("Hectometres")
        cboHabUnits.Items.Add("Dekametres")

        For i As Integer = 0 To cboHabUnits.Items.Count - 1
            If cboHabUnits.Items.Item(i).ToString = m_LayerToAdd.HabUnits Then
                cboHabUnits.SelectedIndex = i
            End If
        Next

    End Sub


    Private Sub cmdSaveLineHabSettings_Click_1(sender As Object, e As EventArgs) Handles cmdSaveLineHabSettings.Click
        ' Update LayerToAddObject

        Dim iIndex As Integer = 0

        Try
            iIndex = cboLengthField.SelectedIndex
            If iIndex <> -1 Then
                m_LayerToAdd.LengthField = cboLengthField.SelectedItem.ToString
            Else
                MsgBox("Please select a length field or click 'cancel'.")
                Exit Sub
            End If
        Catch ex As Exception
            MsgBox("Error on save.  Please select a length field or hit 'cancel' " + ex.Message)
            Exit Sub
        End Try

        Try
            iIndex = cboLengthUnits.SelectedIndex
            If iIndex <> -1 Then
                m_LayerToAdd.LengthUnits = cboLengthUnits.SelectedItem.ToString
            Else
                MsgBox("Please set the length units or click 'cancel'.")
                Exit Sub
            End If
        Catch ex As Exception
            MsgBox("Error on save.  Please set the length units or hit 'cancel' " + ex.Message)
            Exit Sub
        End Try

        Try
            iIndex = cboHabClassField.SelectedIndex
            If iIndex <> -1 Then
                m_LayerToAdd.HabClsField = cboHabClassField.SelectedItem.ToString
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
            iIndex = cboHabUnits.SelectedIndex
            If iIndex <> -1 Then
                m_LayerToAdd.HabUnits = cboHabUnits.SelectedItem.ToString
            Else
                MsgBox("Please set the habitat units or click 'cancel'.")
                Exit Sub
            End If
        Catch ex As Exception
            MsgBox("Error on save.  Please set the habitat units or hit 'cancel' " + ex.Message)
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
        Public Function CompareNames(ByVal obj As LineLayerToAdd) As Boolean
            Return (_name = obj.Layer)
        End Function
    End Class

    Private Sub cmdCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdCancel.Click
        Me.Close()
    End Sub
End Class