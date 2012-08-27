Public Class DelJobRules

    Private _displayMessages As List(Of String)
    Private _rxRanges As Ranges
    Private _rxDefaults As Defaults

    Public Property DisplayMessages() As List(Of String)
        Get
            Return Me._displayMessages
        End Get
        Set(ByVal value As List(Of String))
            Me._displayMessages = value
        End Set
    End Property

    Public Property RxRanges() As Ranges
        Get
            Return Me._rxRanges
        End Get
        Set(ByVal value As Ranges)
            Me._rxRanges = value
        End Set
    End Property

    Public Property RxDefaults() As Defaults
        Get
            Return Me._rxDefaults
        End Get
        Set(ByVal value As Defaults)
            Me._rxDefaults = value
        End Set
    End Property

    Public Class Defaults
        Public FittingVertex As Integer
        Public RefractiveVertex As Integer
        Public PantoTilt As Integer
        Public WrapAngle As Integer
        Public WorkingDistancesDistance As Decimal
        Public WorkingDistancesViewingAngle As String
        Public WorkingDistancesActivity As String
        Public WorkingDistancesNear As Decimal
        Public WorkingDistancesFar As Decimal
    End Class

    Public Class Ranges

        ' Distance RX (R+L)
        Public DistanceSph As New Range
        Public DistanceCyl As New Range
        Public DistanceAxis As New Range

        ' Near RX (R+L)
        Public NearAdd As New Range

        ' Prism (R+L)
        Public PrismIn As New Range
        Public PrismUp As New Range

        ' Vertext/Panto (R+L)
        Public FittingVertex As New Range
        Public RefractiveVertex As New Range
        Public PantoTilt As New Range
        Public WrapAngle As New Range

        ' Fitting Measurements (R+L)
        Public MonocularPD As New Range
        Public FittingHeight As New Range

        ' Frame Data
        Public FrameA As New Range
        Public FrameB As New Range
        Public FrameDbl As New Range
        Public FrameED As New Range

    End Class

    Public Class Range
        Public MinValue As Decimal
        Public MaxValue As Decimal

        Public Sub SetRange(ByVal minValue As Decimal, ByVal maxValue As Decimal)
            If minValue > maxValue Then
                Throw New Exception(String.Format("SetRange method received invalid min/max value range {0}/{1}", minValue, maxValue))
            End If
            Me.MinValue = minValue
            Me.MaxValue = maxValue
        End Sub
    End Class

End Class

