Public Class Menu

    Public Shared Sub Main()
        SingleInstance.Run(New Menu())
    End Sub

    Private Sub LinkLabel1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles LinkLabel1.Click
        Dim f As New CycleCount()
        f.ShowDialog()
    End Sub

    Private Sub LinkLabel2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles LinkLabel2.Click
        Dim f As New Receivings()
        f.ShowDialog()
    End Sub

    Private Sub LinkLabel3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles LinkLabel3.Click
        Dim f As New DELConsolidation()
        f.ShowDialog()
    End Sub

    Private Sub Menu_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        SingleInstance.HideTaskBar(Me.Handle)
    End Sub

    Private Sub Menu_Closed(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Closed
        Me.BringToFront()
        SingleInstance.ShowTaskBar(Me.Handle)
    End Sub


    Private Sub Menu_Closing(ByVal sender As System.Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles MyBase.Closing
        'Ask for PASSWORD
        If InputForm.GetInput("Enter password:", "Password required") <> "odg" Then
            e.Cancel = True
        End If
    End Sub

    Private Sub Menu_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles MyBase.KeyPress
        Select Case e.KeyChar
            Case "1"
                LinkLabel1_Click(Me, Nothing)
            Case "2"
                LinkLabel2_Click(Me, Nothing)
            Case "3"
                LinkLabel3_Click(Me, Nothing)
        End Select
    End Sub

End Class
