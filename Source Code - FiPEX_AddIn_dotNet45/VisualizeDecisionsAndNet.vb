Public Class VisualizeDecisionsAndNet
    Inherits ESRI.ArcGIS.Desktop.AddIns.Button
    Private m_FiPEx__1 As FishPassageExtension

  Public Sub New()

  End Sub

  Protected Overrides Sub OnClick()
        My.ArcMap.Application.CurrentTool = Nothing
        If m_FiPEx__1 Is Nothing Then
            m_FiPEx__1 = FiPEX_ArcMap_10p4_up_AddIn_dotNet45_2020.FishPassageExtension.GetExtension()
        End If
        Using MyForm As New FiPEX_ArcMap_10p4_up_AddIn_dotNet45_2020.frmVisualizeDecisionsAndNet
            If MyForm.Form_Initialize(My.ArcMap.Application) Then
                MyForm.ShowDialog()
            End If
        End Using
    End Sub

    Protected Overrides Sub OnUpdate()
        If m_FiPEx__1 Is Nothing Then
            m_FiPEx__1 = FiPEX_ArcMap_10p4_up_AddIn_dotNet45_2020.FishPassageExtension.GetExtension()
        End If
    End Sub
End Class
