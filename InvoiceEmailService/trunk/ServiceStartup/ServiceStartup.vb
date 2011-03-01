Imports InvoiceEmail
Imports ServiceShared
Imports System.Windows.Forms
Imports System.IO

Public Class ServiceStartup

    Private _config As ServiceConfig

    ' add services as required in conjunction with adding classes to ServiceEngine project

    Private Const ConfigFileName As String = "OrdersImportService.xml"
    Private _configFileName As String = _
    Application.StartupPath & "\" & ConfigFileName
    Private _engine As InvoiceEmail.ServiceEngine

    Protected Overrides Sub OnStart(ByVal args() As String)
        ' Add code here to start your service. This method should set things
        ' in motion so your service can do its work.

        Me.WriteToLog(0, "OnStart")
        ' also write to event log

        Try
            Me._config = XMLHelper.DeSerializeXml( _
            GetType(ServiceConfig), _
            _configFileName)

        Catch ex As Exception
            Me._config = New ServiceConfig

        End Try

        'Me.WriteToLog(0, "Listening on Port " & Me._config.Port.ToString)

        Try
            Me._engine = InvoiceEmail.ServiceEngine.GetInstance(Me._config)

            AddHandler _engine.StatusUpdate, AddressOf StatusUpdate
            '        AddHandler cmbctl.KeyDown, AddressOf cmb_KeyDown
            Me._engine.Start()

        Catch ex As Exception
            Me.WriteToLog(0, "Fatal Error Starting Engine (" & ex.Message & ")")
            ' also write to event log
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

        'lenni thinks that this approach will not solve the multi-threading problem


    End Sub

    'Private Sub RecordMessage(ByVal e As StatusUpdateEventArgs)
    '    Me.lstStatus.Items.Add(e.Message & " on thread " & e.ThreadId)
    'End Sub

    Protected Overrides Sub OnStop()
        ' Add code here to perform any tear-down necessary to stop your service.
        Me.WriteToLog(0, "OnStop ... Waiting for Engine to Stop")
        Me.WriteToLog(0, "OnStop ... Engine has Stopped")
    End Sub

End Class