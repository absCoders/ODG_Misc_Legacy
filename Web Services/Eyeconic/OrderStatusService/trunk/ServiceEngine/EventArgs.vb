Public Delegate Sub StatusUpdateEventHandler( _
ByVal sender As Object, _
ByVal e As StatusUpdateEventArgs)

Public Class StatusUpdateEventArgs
    Inherits EventArgs
    Private _message As String
    Private _threadId As Integer

    ''' <summary>
    ''' This is the custom message
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property Message() As String
        Get
            Return Me._message
        End Get
    End Property

    Friend Sub New(ByVal message As String, ByVal threadId As Integer)
        Me._message = message
        Me._threadId = threadId
    End Sub

    Public ReadOnly Property ThreadId() As Integer
        Get
            Return Me._threadId
        End Get
    End Property
End Class