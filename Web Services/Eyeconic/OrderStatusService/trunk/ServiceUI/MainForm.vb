Imports ServiceEngine
Imports ServiceShared

Public Class MainForm

    Private _config As ServiceConfig
    Private _engine As ServiceEngine.ServiceEngine
    Private Const ConfigFileName As String = "service.xml"
    Private _configFileName As String = _
    Application.StartupPath & "\" & ConfigFileName

    Protected Overrides Sub OnLoad( _
    ByVal e As System.EventArgs)

        ' my stuff before their stuff

        MyBase.OnLoad(e)

        Try
            Me._config = XMLHelper.DeSerializeXml( _
            GetType(ServiceConfig), _
            _configFileName)
        Catch ex As Exception
            Me._config = New ServiceConfig
        End Try

        Me.PropertyGrid1.SelectedObject = Me._config

        ' my stuff after their stuff, 
        ' usually my code will go here

    End Sub

    Protected Overrides Sub OnClosed( _
    ByVal e As System.EventArgs)
        MyBase.OnClosed(e)
        ' serialize the configuration
        XMLHelper.SerializeXml(Me._config, Me._configFileName)

    End Sub

    Private Sub cmdRun_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdRun.Click
        Me.lstStatus.Items.Clear()

        Me._engine = ServiceEngine.ServiceEngine.GetInstance(Me._config)

        AddHandler _engine.StatusUpdate, AddressOf StatusUpdate

        Me._engine.Start()
    End Sub

    Private Sub StatusUpdate(ByVal sender As Object, ByVal e As StatusUpdateEventArgs)

        If e.ThreadId = 0 Then
            Me.RecordMessage(e)
        Else
            Console.WriteLine(e.Message)
        End If

    End Sub

    Private Sub RecordMessage(ByVal e As StatusUpdateEventArgs)
        Me.lstStatus.Items.Add(e.Message & " on thread " & e.ThreadId)
    End Sub
End Class