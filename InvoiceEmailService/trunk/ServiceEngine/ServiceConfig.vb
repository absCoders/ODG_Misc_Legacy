Imports System.ComponentModel
Imports System.Xml

Public Class ServiceConfig

    ' IMPORTANT - remember to change the following of ServiceStartup project
    ' 1) Assembly Name (in Properties)
    ' 2) ConfigFileName in ServiceStartup.vb
    ' 3) ServiceName, DisplayName and Description in ServiceInstallation.vb

    Private Const DefaultTNS As String = "TST"

    Private _TNS As String = DefaultTNS
    Private _UID As String = DefaultTNS
    Private _PWD As String = DefaultTNS

    Private _FileFolder As String = String.Empty
    Private _StartEmail As String = "2200"
    Private _EmailDay As String = "ALL"


    Public Sub New()
        'Get settings from folder

        _FileFolder = My.Application.Info.DirectoryPath
        If _FileFolder.EndsWith("\") Then _FileFolder &= "\"

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
                                Case "StartEmail"
                                    _StartEmail = xReader.ReadElementContentAsString()
                                Case "EmailDay"
                                    _EmailDay = xReader.ReadElementContentAsString()
                            End Select
                    End Select
                Loop
                xReader.Close()
            End Using
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

    <DefaultValue(DefaultTNS)> _
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

    <DefaultValue(DefaultTNS)> _
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

    <DefaultValue("2200")> _
    <Category("EmailInvoice")> _
    <Description("This is the time (military) to start sending todays invoices")> _
    <DisplayName("Time to start emailing invoices to customers")> _
    Public Property StartEmailing() As String
        Get
            Return Me._StartEmail
        End Get
        Set(ByVal value As String)
            Me._StartEmail = value
        End Set
    End Property

    <DefaultValue("Application Directory")> _
    <Category("EmailInvoice")> _
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

    <DefaultValue("ALL")> _
    <Category("EmailInvoice")> _
    <Description("This is the Day to email invoices")> _
    <DisplayName("Day to email invoices to customers. ALL for everyday, else first three chars of day")> _
    Public Property EmailDay() As String
        Get
            Return _EmailDay
        End Get
        Set(ByVal value As String)
            _EmailDay = (value & String.Empty).ToString.ToUpper.Trim
        End Set
    End Property
End Class