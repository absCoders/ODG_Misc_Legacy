Public Class InputForm

    Public Shared Function GetInput(ByVal message As String, ByVal caption As String) As String
        Dim x As New InputForm()
        x.Text = caption
        x.lblText.Text = message
        x.ShowDialog()
        Return x.txtInput.Text
    End Function

    Private Sub btnOK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnOK.Click
        Me.Close()
    End Sub

    Private Sub txtInput_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtInput.KeyPress
        If e.KeyChar = Chr(13) Then 'Chr(13) is the Enter Key
            Me.Close()
        End If
    End Sub
End Class