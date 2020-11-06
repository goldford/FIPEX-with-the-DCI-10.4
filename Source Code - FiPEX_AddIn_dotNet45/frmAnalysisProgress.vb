'Imports System.Runtime.InteropServices
'Imports System.Drawing
'Imports ESRI.ArcGIS.ADF.BaseClasses
'Imports ESRI.ArcGIS.ADF.CATIDs
'Imports ESRI.ArcGIS.Framework
'Imports ESRI.ArcGIS.ArcMapUI
'Imports ESRI.ArcGIS.Display
'Imports ESRI.ArcGIS.esriSystem
'Imports ESRI.ArcGIS.EditorExt
'Imports ESRI.ArcGIS.Geoprocessing
'Imports ESRI.ArcGIS.Geodatabase
'Imports ESRI.ArcGIS.Geometry
'Imports System.Text.RegularExpressions
'Imports ESRI.ArcGIS.DataSourcesGDB
'Imports ESRI.ArcGIS.NetworkAnalysis
'Imports ESRI.ArcGIS.Carto
'Imports ESRI.ArcGIS.DataSourcesOleDB ' For DCI calculation
'Imports ESRI.ArcGIS.GeoDatabaseUI    ' For use with DCI table conversion - IExportInterface
'Imports System.IO                    ' For reading DCI output file
'Imports System                       ' added based on MSDN instructions for writing test txt file (for permission check)
'Imports System.ComponentModel
'Imports System.Threading


Public Class frmAnalysisProgress
    'Private m_MxDoc As IMxDocument
    'Private m_app As ESRI.ArcGIS.Framework.IApplication
    'Private Shared m_UNAExt_2 As IUtilityNetworkAnalysisExt
    'Private Shared m_pNetworkAnalysisExt_2 As INetworkAnalysisExt
    'Private Shared m_FiPEx__2 As FishPassageExtension
    Friend m_bCloseMe As Boolean = False

    Private Delegate Sub ChangeLabelCallback(ByVal item As String)
    Private Delegate Sub ChangeProgressBarCallback(ByVal value1 As Integer)
    'Make thread-safe calls to Windows Forms Controls.
    Friend Sub ChangeLabel(ByVal item As String)
        ' InvokeRequired compares the thread ID of the
        'calling thread to the thread ID of the creating thread.
        ' If these threads are different, it returns true.
        ' see here for an explanation of this bullshit
        ' http://edndoc.esri.com/arcobjects/9.2/NET/2c2d2655-a208-4902-bf4d-b37a1de120de.htm
        If Me.lblProgress.InvokeRequired Then
            'Call itself on the main thread.
            Dim d As New ChangeLabelCallback(AddressOf ChangeLabel)
            Me.Invoke(d, New Object() {item})
        Else
            'Guaranteed to run on the main UI thread. 
            Me.lblProgress.Text = item
        End If
    End Sub

    'Make thread-safe calls to Windows Forms Controls.
    Friend Sub ChangeProgressBar(ByVal value1 As Integer)
        ' InvokeRequired compares the thread ID of the
        'calling thread to the thread ID of the creating thread.
        ' If these threads are different, it returns true.
        ' see here for an explanation of this bullshit
        'http://edndoc.esri.com/arcobjects/9.2/NET/2c2d2655-a208-4902-bf4d-b37a1de120de.htm
        If Me.ProgressBar1.InvokeRequired Then
            'Call itself on the main thread.
            Dim d As New ChangeProgressBarCallback(AddressOf ChangeProgressBar)
            Me.Invoke(d, New Object() {value1})
        Else
            'Guaranteed to run on the main UI thread. 
            Me.ProgressBar1.Value = value1

        End If

        If value1 = 100 Then
            m_bCloseMe = True
            Me.Close()
        End If
    End Sub

    Private Sub frmAnalysisProgress_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load


    End Sub
    'Public Function Form_Initialize(ByVal m_application As ESRI.ArcGIS.Framework.IApplication) As Boolean
    Public Function Form_Initialize() As Boolean
        Try
            'm_app = m_application
            'm_MxDoc = CType(m_app.Document, ESRI.ArcGIS.ArcMapUI.IMxDocument)
            'm_FiPEx__2 = FishPassageExtension.GetExtension
            'm_UNAExt_2 = FishPassageExtension.GetUNAExt

            Return True

        Catch ex As Exception
            System.Windows.Forms.MessageBox.Show(ex.ToString, "Initialize")
            Return False
        End Try
    End Function

  
    Private Sub cmdCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdCancel.Click
        m_bCloseMe = True
        Me.Close()
    End Sub

    Private Sub frmAnalysisProgress_Closing(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.FormClosing
        ''Windows.Forms.Control.CheckForIllegalCrossThreadCalls = False
        'BackgroundWorker1.WorkerSupportsCancellation = True

        'If Not BackgroundWorker1.IsBusy = True Then
        '    BackgroundWorker1.RunWorkerAsync(m_UNAExt_2)
        'End If
        m_bCloseMe = True
        Me.Dispose()
    End Sub

End Class