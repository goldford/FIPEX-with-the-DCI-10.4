Public Class frmBulkSnap
    Public m_bCancel As Boolean
    Private Sub cmdOk_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdOK.Click
        m_bCancel = False
        Me.Close()
    End Sub

    Private Sub cmdCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdCancel.Click
        m_bCancel = True
    End Sub
End Class