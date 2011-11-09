Imports ServiceEngine
Imports ServiceShared
Imports System.Windows.Forms
Imports System.IO

Public Class ServiceStartup

    Private _config As ServiceConfig
    Private _engine As ServiceEngine.ServiceEngine
    Private Const ConfigFileName As String = "OrderStatusService.xml"
    Private _configFileName As String = _
    Application.StartupPath & "\" & ConfigFileName

    Protected Overrides Sub OnStart(ByVal args() As String)
        ' Add code here to start your service. This method should set things
        ' in motion so your service can do its work.

        Me.WriteToLog(0, "OnStart")
        ' also write to event log

        Try
            Me._config = XMLHelper.DeSerializeXml(GetType(ServiceConfig),configFileName)
        Catch ex As Exception
            Me._config = New ServiceConfig
        End Try

        'Config values
        Me.WriteToLog(0, "Using config value: " & Me._config.ConfigValue.ToString)

        Try
            Me._engine = ServiceEngine.ServiceEngine.GetInstance(Me._config)

            AddHandler _engine.StatusUpdate, AddressOf StatusUpdate

            Me._engine.Start()

        Catch ex As Exception
            Me.WriteToLog(0, "Fatal Error Starting Engine (" & ex.Message & ")")
        End Try

    End Sub

    Sub WriteToLog(ByVal threadId As Integer, ByVal message As String)

        Dim filename As String = Application.StartupPath _
         & "\logfile" & threadId.ToString & ".txt"

        ' filename should have a date/time stamp

        Dim sw As New StreamWriter(filename, True)
        sw.WriteLine(message)
        sw.Close()

        ' write to event log
    End Sub

    Private Sub StatusUpdate(ByVal sender As Object, _
    ByVal e As StatusUpdateEventArgs)
        Me.WriteToLog(e.ThreadId, e.Message)
    End Sub


    Protected Overrides Sub OnStop()
        ' Add code here to perform any tear-down necessary to stop your service.
        Me.WriteToLog(0, "OnStop ... Waiting for Engine to Stop")
        Me._engine.Cancel()
        Me.WriteToLog(0, "OnStop ... Engine has Stopped")
    End Sub

End Class