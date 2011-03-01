Imports System.Threading
Imports InvoiceEmail.Extensions

Public Class ServiceEngine
    Private _config As ServiceConfig
    Private Shared _instance As ServiceEngine
    Private _isRunning As Boolean = False
    Private _backgroundThread As Thread

    Public Event StatusUpdate As StatusUpdateEventHandler

    Private emailer As New InoviceEmail.InvoiceEmailer

    'Public Sub New()
    '    MyBase.New()
    'End Sub

    Public Shared Function GetInstance(ByVal config As ServiceConfig) As ServiceEngine

        If _instance Is Nothing Then
            _instance = New ServiceEngine(config)
        End If

        Return _instance

    End Function

    Private Sub New(ByVal config As ServiceConfig)
        'Public Sub New(ByVal config As ServiceConfig)
        ' make the sub Private to implement as a singleton 
        ' - you need Public Shared Function above

        MyBase.New() ' this line is implied in VB, and not necessary.  
        ' It must be the 1st line
        ' nothing may be before it

        If config Is Nothing Then
            Throw New Exception("Config cannot be Nothing")
            Return ' no return statement is necessary
        End If

        Me._config = config

    End Sub

    Public Sub Start()
        If Me._isRunning Then
            Throw New Exception("Service is already started")
        End If

        Me._isRunning = True

        RaiseEvent StatusUpdate(Me, New StatusUpdateEventArgs("Service is Started", 0))

        'Dim thread As New Thread(AddressOf StartListening)
        Me._backgroundThread = New Thread(AddressOf StartListening)
        Me._backgroundThread.Start()

    End Sub

    Private Sub StartListening()

        RaiseEvent StatusUpdate(Me, New StatusUpdateEventArgs("Service is Listening", 1))

        Try
            Start_Service()

        Catch ex As Exception
            RaiseEvent StatusUpdate(Me, New StatusUpdateEventArgs("Unexpected Error (" & ex.Message & ")", 1))

        End Try

    End Sub

    Public Sub Cancel()

        If Me._backgroundThread Is Nothing Then
            Return
        End If
        Stop_Service()
        Me._backgroundThread.Abort()

    End Sub

    Private Sub Stop_Service()
        emailer.CloseLog()
    End Sub

    Private Sub Start_Service()

        With My.Computer.FileSystem
            If Not .DirectoryExists(Me._config.FileFolder & "LOGS\") Then
                .CreateDirectory(Me._config.FileFolder & "LOGS\")
            End If

            If Not .DirectoryExists(Me._config.FileFolder & "ARCHIVE\") Then
                .CreateDirectory(Me._config.FileFolder & "ARCHIVE\")
            End If
        End With

        emailer.LogIn()
    End Sub

    Private Delegate Sub DataReceivedDelegate( _
    ByVal dtResponses As DataTable, _
    ByVal DATA_RECEIVED As String, _
    ByVal FILENAME As String)

    Public Sub DataReceived( _
    ByVal dtResponses As DataTable, _
    ByVal DATA_RECEIVED As String, _
    ByVal FILENAME As String)

    End Sub
End Class