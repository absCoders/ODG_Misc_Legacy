Imports System.ComponentModel

Public Class ServiceConfig

    ' IMPORTANT - remember to change the following of ServiceStartup project
    ' 1) Assembly Name (in Properties)
    ' 2) ConfigFileName in ServiceStartup.vb
    ' 3) ServiceName, DisplayName and Description in ServiceInstallation.vb

    Private Const DefaultTNS = "TST"
    Private _TNS As String = DefaultTNS

    Private Const DefaultUID = "TST"
    Private _UID As String = DefaultUID

    Private Const DefaultPWD = "TST"
    Private _PWD As String = DefaultPWD

    Private Const DefaultFileFolder = "C:\Orders\Logs\"
    Private _FileFolder As String = DefaultFileFolder

    Private _created As Date
    Private _backcolor As System.Drawing.Color

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

End Class