Public Class Logger

    Private logStreamWriter As System.IO.StreamWriter

    Public Sub New(ByVal logDirectory As String)
        OpenLogFile(logDirectory)
    End Sub


    Public Function OpenLogFile(ByVal logDirectory As String) As Boolean

        Try

            If Not My.Computer.FileSystem.DirectoryExists(logDirectory) Then
                My.Computer.FileSystem.CreateDirectory(logDirectory)
            End If

            Dim logFilename As String = Format(Now, "yyyyMMdd") & ".log"
            If logStreamWriter IsNot Nothing Then
                logStreamWriter.Close()
                logStreamWriter.Dispose()
            End If

            If Not logdirectory.EndsWith("\") Then logdirectory &= "\"
            logdirectory &= "Logs\"
            If Not My.Computer.FileSystem.DirectoryExists(logdirectory) Then
                My.Computer.FileSystem.CreateDirectory(logdirectory)
            End If

            logStreamWriter = New System.IO.StreamWriter(logDirectory & logFilename, True)

            Return True

        Catch ex As Exception
            Return False
        End Try
    End Function

    Public Sub RecordLogEntry(ByVal message As String)
        Try
            logStreamWriter.WriteLine(DateTime.Now & ": " & message)
        Catch ex As Exception
            'Failed to record log entry -- handle this how?
        End Try
    End Sub

    Public Sub CloseLog()
        If logStreamWriter IsNot Nothing Then
            logStreamWriter.Close()
            logStreamWriter.Dispose()
            logStreamWriter = Nothing
        End If
    End Sub

End Class

