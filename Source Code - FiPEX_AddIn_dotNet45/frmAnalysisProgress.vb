
Public Class frmAnalysisProgress

    Friend m_bCloseMe As Boolean = False
    Private Delegate Sub ChangeLabelCallback(ByVal item As String)
    Private Delegate Sub ChangeProgressBarCallback(ByVal value1 As Integer)

    'Make thread-safe calls to Windows Forms Controls.
    Friend Sub ChangeLabel(ByVal item As String)
        ' InvokeRequired compares the thread ID of the calling thread to the thread ID of the creating thread.
        ' If these threads are different, it returns true. see here for an explanation of this bullshit
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
        ' InvokeRequired compares the thread ID of the calling thread to the thread ID of the creating thread.
        ' If these threads are different, it returns true. see here for an explanation of this bullshit
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
    Public Function Form_Initialize() As Boolean
        Try
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