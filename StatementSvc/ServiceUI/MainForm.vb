Imports StatementsEmail
Imports ServiceShared
Imports System.IO

Public Class MainForm

    Private _config As ServiceConfig
    Private _engine As StatementsEmail.ServiceEngine
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
        ' serialize the configuation
        XMLHelper.SerializeXml(Me._config, Me._configFileName)

    End Sub

    'Private Sub MainForm_Load( _
    'ByVal sender As Object, _
    'ByVal e As System.EventArgs) _
    'Handles Me.Load

    'End Sub

    Private Sub cmdStartService_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdStartService.Click
        Me.lstStatus.Items.Clear()

        Me._engine = StatementsEmail.ServiceEngine.GetInstance(Me._config)

        AddHandler _engine.StatusUpdate, AddressOf StatusUpdate
        '        AddHandler cmbctl.KeyDown, AddressOf cmb_KeyDown
        Me._engine.Start()

        cmdStartService.Enabled = False
        cmdStopService.Enabled = True

    End Sub

    Private Sub StatusUpdate(ByVal sender As Object, ByVal e As StatusUpdateEventArgs)

        If e.ThreadId = 0 Then
            Me.RecordMessage(e)
        Else
            'Me.RecordMessage(e)
            ' homework
            Me.Invoke(New RecordMessageDelegate(AddressOf RecordMessage), e)
            Console.WriteLine(e.Message)
            'Me.WriteToLog(e.ThreadId, e.Message)

        End If

        '    Me.Invoke(New _
        '   DataReceivedDelegate(AddressOf DataReceived), _
        '   PORT_ID, DATA_RECEIVED)
        'End Sub

        'Private Delegate Sub DataReceivedDelegate( _
        'ByVal PORT_ID As String, _
        'ByVal DATA_RECEIVED As String)

    End Sub

    Private Delegate Sub RecordMessageDelegate(ByVal e As StatusUpdateEventArgs)

    Private Sub RecordMessage(ByVal e As StatusUpdateEventArgs)
        Me.lstStatus.Items.Add(e.Message & " on thread " & e.ThreadId)
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

    Private Sub cmdStopService_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdStopService.Click
        Me._engine = StatementsEmail.ServiceEngine.GetInstance(Me._config)

        AddHandler _engine.StatusUpdate, AddressOf StatusUpdate
        '        AddHandler cmbctl.KeyDown, AddressOf cmb_KeyDown
        Me._engine.Cancel()

        cmdStartService.Enabled = True
        cmdStopService.Enabled = False
    End Sub
End Class