Imports System.Threading

Public Class ServiceEngine

    Private _config As ServiceConfig
    Private Shared _instance As ServiceEngine
    Private _isRunning As Boolean = False
    Private _backgroundThread As Thread

    Private serviceTimer As Timer



    Private filefolder As String

    Public Event StatusUpdate As StatusUpdateEventHandler


    Public Shared Function GetInstance(ByVal config As ServiceConfig) As ServiceEngine

        If _instance Is Nothing Then
            _instance = New ServiceEngine(config)
        End If

        Return _instance

    End Function

    Private Sub New(ByVal config As ServiceConfig)
        MyBase.New()

        If config Is Nothing Then
            Throw New Exception("Config cannot be nothing")
            Return ' no return statement is necessary
        End If

        Me._config = config

    End Sub

    Public Sub Start()
        If Me._isRunning Then
            Throw New Exception("Service is already started")
        End If

        Me._isRunning = True

        Dim serviceTimerCallback As New TimerCallback(AddressOf DoOrderStatusUpdates)
        serviceTimer = New Timer(serviceTimerCallback, Nothing, 1000, 1000 * 60 * 5)

        RaiseEvent StatusUpdate(Me, New StatusUpdateEventArgs("Service is Started", 0))
    End Sub


    Public Sub Cancel()
        serviceTimer.Change(Timeout.Infinite, Timeout.Infinite)
        serviceTimer.Dispose()
    End Sub

    Private Sub DoOrderStatusUpdates() 'Each call of this function happens in its own thread

        Try
            Using updater As New orderStatusUpdater()
                If updater.loadedSuccessfully Then
                    updater.DoOrderStatusUpdates()
                End If
            End Using
        Catch ex As Exception

        End Try

    End Sub



End Class