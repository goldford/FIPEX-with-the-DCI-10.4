Imports ESRI.ArcGIS.ArcMapUI
Public Class frmProgress_DistanceToMouth
    Public m_iCurrentFeature
    Public m_iTotalFeatures
    Private m_MxDoc As IMxDocument
    Public Function Form_Initialize(ByVal m_application As ESRI.ArcGIS.Framework.IApplication) As Boolean
        Try

            m_MxDoc = My.ArcMap.Application.Document
            'm_FiPEx = FishPassageExtension.GetExtension

            Return True

        Catch ex As Exception
            System.Windows.Forms.MessageBox.Show(ex.ToString, "Initialize")
            Return False
        End Try
    End Function
End Class