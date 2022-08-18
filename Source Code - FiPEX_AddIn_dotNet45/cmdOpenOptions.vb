Public Class cmdOpenOptions
    Inherits ESRI.ArcGIS.Desktop.AddIns.Button
    Private m_FiPEx__1 As FishPassageExtension
    Private m_UNAExt As ESRI.ArcGIS.EditorExt.IUtilityNetworkAnalysisExt
    Private m_pNetworkAnalysisExt As ESRI.ArcGIS.EditorExt.INetworkAnalysisExt

  Public Sub New()

  End Sub

  Protected Overrides Sub OnClick()
    '
    '  TODO: Sample code showing how to access button host
    '
        My.ArcMap.Application.CurrentTool = Nothing
        Dim FiPEx__1 As FishPassageExtension = FishPassageExtension.GetExtension
        '2020 - switching out forms, passing info about how the user accessed form
        '     so that the 'run button' etc associated with 'advanced analysis' can be
        '     hidden
        'Using MyForm As New FiPEX_ArcMap_10p4_up_AddIn_dotNet45_2022.Options
        'If MyForm.Form_Initialize(My.ArcMap.Application) Then
        'MyForm.ShowDialog()
        'End If
        'End Using
        Using MyForm As New FiPEX_ArcMap_10p4_up_AddIn_dotNet45_2022.frmRunAdvancedAnalysis
            If MyForm.Form_Initialize(My.ArcMap.Application) Then
                MyForm.m_FIPEXOptionsContext = "optionsmenu"
                MyForm.ShowDialog()
            End If
        End Using
    End Sub

    Protected Overrides Sub OnUpdate()
        If My.ArcMap.Application IsNot Nothing Then

            Dim FiPEx__1 As FishPassageExtension = FishPassageExtension.GetExtension

            If m_FiPEx__1 Is Nothing Then
                m_FiPEx__1 = FiPEX_ArcMap_10p4_up_AddIn_dotNet45_2022.FishPassageExtension.GetExtension()
            End If
            If m_UNAExt Is Nothing Then
                m_UNAExt = FiPEX_ArcMap_10p4_up_AddIn_dotNet45_2022.FishPassageExtension.GetUNAExt
            End If
            'Dim FiPEx__1 As FishPassageExtension = FiPEX_AddIn_dotNet35_2.FishPassageExtension.GetExtension
            'Dim pUNAExt As IUtilityNetworkAnalysisExt = FiPEX_AddIn_dotNet35_2.FishPassageExtension.GetUNAExt
            If m_pNetworkAnalysisExt Is Nothing Then
                m_pNetworkAnalysisExt = CType(m_UNAExt, ESRI.ArcGIS.EditorExt.INetworkAnalysisExt)
            End If

            If m_pNetworkAnalysisExt.NetworkCount > 0 Then
                Me.Enabled = True
            Else
                Me.Enabled = False
            End If


            'If FiPEx__1.m_enableState = ESRI.ArcGIS.Desktop.AddIns.ExtensionState.Enabled Then
            '    Enabled = FiPEx__1.HasNetwork
            'Else
            '    Enabled = False
            'End If
        End If
    End Sub
End Class
