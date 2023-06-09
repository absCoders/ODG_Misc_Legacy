Imports System.ComponentModel
Imports System.Xml

Public Class ServiceConfig

    ' IMPORTANT - remember to change the following of ServiceStartup project
    ' 1) Assembly Name (in Properties)
    ' 2) ConfigFileName in ServiceStartup.vb
    ' 3) ServiceName, DisplayName and Description in ServiceInstallation.vb

    Private Const DefaultTNS = "ODG"
    Private _TNS As String = DefaultTNS

    Private Const DefaultUID = "ODG"
    Private _UID As String = DefaultUID

    Private Const DefaultPWD = "ODG"
    Private _PWD As String = DefaultPWD

    Private Const DefaultFileFolder = "C:\OrdersImport\" & DefaultTNS & "\"
    Private _FileFolder As String = DefaultFileFolder

    Private _created As Date
    Private _backcolor As System.Drawing.Color

    Private _period As Integer = 10
    Private _driveLetter As String = String.Empty
    Private _driveLetterIP As String = String.Empty


    Public Sub New()
        'Get settings from folder

        If My.Computer.FileSystem.FileExists(My.Application.Info.DirectoryPath & "\SvcConfig.xml") Then
            Using xReader As Xml.XmlTextReader = New XmlTextReader(My.Application.Info.DirectoryPath & "\SvcConfig.xml")
                Do While xReader.Read()
                    Select Case xReader.NodeType
                        Case XmlNodeType.Element
                            Select Case xReader.Name
                                Case "DefaultTNS"
                                    _TNS = xReader.ReadElementContentAsString()
                                Case "DefaultUID"
                                    _UID = xReader.ReadElementContentAsString()
                                Case "DefaultPWD"
                                    _PWD = xReader.ReadElementContentAsString()
                                Case "DefaultFileFolder"
                                    _FileFolder = xReader.ReadElementContentAsString()
                                Case "DefaultFolder"
                                    _FileFolder = My.Application.Info.DirectoryPath & "\" ' & xReader.ReadElementContentAsString() & "\
                                Case "DefaultPeriod"
                                    TimerPeriod = Convert.ToInt16(xReader.ReadElementContentAsString())
                                Case "DriveLetter"
                                    _driveLetter = xReader.ReadElementContentAsString()
                                Case "DriveLetterIP"
                                    _driveLetterIP = xReader.ReadElementContentAsString()
                            End Select
                    End Select
                Loop
                xReader.Close()
            End Using
        Else
            _FileFolder = My.Application.Info.DirectoryPath & "\"
        End If

    End Sub

    <DefaultValue(DefaultTNS)> _
    <Category("Oracle")> _
    <Description("This is the TNS Name of the Oracle Database that the Service write results to")> _
    <DisplayName("TNS Name of Oracle Schema to Log Results to")> _
    Public Property TNS() As String
        Get
            Return Me._TNS
        End Get
        Set(ByVal value As String)
            Me._TNS = value
        End Set
    End Property

    <DefaultValue(DefaultUID)> _
    <Category("Oracle")> _
    <Description("This is the UID of the Oracle Database that the Service write results to")> _
    <DisplayName("UID of Oracle Schema to Log Results to")> _
    Public Property UID() As String
        Get
            Return Me._UID
        End Get
        Set(ByVal value As String)
            Me._UID = value
        End Set
    End Property

    <DefaultValue(DefaultPWD)> _
    <Category("Oracle")> _
    <Description("This is the PWD of the Oracle Schema UID")> _
    <DisplayName("PWD of Oracle Schema UID")> _
    Public Property PWD() As String
        Get
            Return Me._PWD
        End Get
        Set(ByVal value As String)
            Me._PWD = value
        End Set
    End Property

    <DefaultValue(DefaultFileFolder)> _
    <Category("OrdersImport")> _
    <Description("This is the Name of the FileFolder where Logs will be stored")> _
    <DisplayName("Name of the FileFolder where Logs will be stored")> _
    Public Property FileFolder() As String
        Get
            Return Me._FileFolder
        End Get
        Set(ByVal value As String)
            Me._FileFolder = value
        End Set
    End Property

    Public Property Created() As Date
        Get
            Return Me._created
        End Get
        Set(ByVal value As Date)
            Me._created = value
        End Set
    End Property

    Public Property BackColor() As System.Drawing.Color
        Get
            Return Me._backcolor
        End Get
        Set(ByVal value As System.Drawing.Color)
            Me._backcolor = value
        End Set
    End Property

    Public Property TimerPeriod() As Integer
        Get
            Return Me._period
        End Get
        Set(ByVal value As Integer)
            Me._period = value
            If _period < 0 Then
                _period = 10
            ElseIf _period > 60 Then
                _period = 60
            End If
        End Set
    End Property

    Public Property DriveLetter()
        Get
            Return _driveLetter
        End Get
        Set(ByVal value)
            _driveLetter = value
        End Set
    End Property

    Public Property DriveLetterIP()
        Get
            Return _driveLetterIP
        End Get
        Set(ByVal value)
            _driveLetterIP = value
        End Set
    End Property

End Class