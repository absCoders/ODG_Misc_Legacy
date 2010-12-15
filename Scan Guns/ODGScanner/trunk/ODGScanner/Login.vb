Public Class Login

    Private Sub txtOperID_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtOperID.KeyPress
        If e.KeyChar = Chr(13) Then
            btnBegin_Click(Me, Nothing)
        End If
    End Sub

    Public Function GetID() As String()
        Me.ShowDialog()
        Return New String() {txtOperID.Text, cmbWarehouse.SelectedItem.ToString}
    End Function

    Private Sub Form2_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        cmbWarehouse.Items.Add("001")
        cmbWarehouse.Items.Add("002")
        cmbWarehouse.Items.Add("003")
        cmbWarehouse.SelectedItem = cmbWarehouse.Items(0)
    End Sub

    Private Sub btnBegin_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnBegin.Click
        If txtOperID.TextLength = 3 Or txtOperID.TextLength = 4 Then
            Me.Close()
        End If
    End Sub

    Private Sub Form2_Closing(ByVal sender As System.Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles MyBase.Closing
        If txtOperID.TextLength <> 3 And txtOperID.TextLength <> 4 Then
            e.Cancel = True
            MsgBox("You must enter an operator ID.")
        End If
    End Sub
End Class