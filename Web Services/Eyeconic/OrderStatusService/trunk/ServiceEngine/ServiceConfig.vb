Imports System.ComponentModel


    Public Class ServiceConfig

        ' IMPORTANT - remember to change the following of ServiceStartup project
        ' 1) Assembly Name (in Properties)
        ' 2) ConfigFileName in ServiceStartup.vb
        ' 3) ServiceName, DisplayName and Description in ServiceInstallation.vb

#Const TEST = True

        Private Const DefaultConfigValue = 1
        Private _ConfigValue As Integer = DefaultConfigValue

#If TEST = True Then
        Private Const DefaultLogDirectory = "C:\Eyeconic\_Test\OrderStatus"
#Else
    Private Const DefaultLogDirectory = "C:\Eyeconic\OrderStatus"
#End If
        Private _logDirectory As String = DefaultLogDirectory

#If TEST = True Then
        Private Const DefaultConnectionString = "Data Source=TST;User ID=TST;Password=TST;pooling=true"
#Else
    Private Const DefaultConnectionString = "Data Source=ODG;User ID=ODG;Password=ODG;pooling=true"
#End If
        Private _connectionString As String = DefaultConnectionString

#If TEST = True Then
        Private Const DefaultServicePassPhrase = "temppassword"
#Else
    Private Const DefaultServicePassPhrase = "3y3c0n1c"
#End If
        Private _servicePassPhrase As String = DefaultServicePassPhrase



        <DefaultValue(DefaultConfigValue)> _
        <Category("ServiceName")> _
        <Description("Description of this config value")> _
        <DisplayName("DisplayName for config value")> _
        Public Property ConfigValue() As Integer
            Get
                Return Me._ConfigValue
            End Get
            Set(ByVal value As Integer)
                Me._ConfigValue = value
            End Set
        End Property


        <DefaultValue(DefaultLogDirectory)> _
        <Category("Logging")> _
        <Description("Directory to save log files to")> _
        <DisplayName("LogDirectory")> _
        Public Property LogDirectory() As String
            Get
                Return Me._logDirectory
            End Get
            Set(ByVal value As String)
                Me._logDirectory = value
            End Set
        End Property

        <DefaultValue(DefaultConnectionString)> _
    <Category("Oracle")> _
    <Description("Connection string for connecting to Oracle")> _
    <DisplayName("ConnectionString")> _
    Public Property ConnectionString() As String
            Get
                Return Me._connectionString
            End Get
            Set(ByVal value As String)
                Me._connectionString = value
            End Set
        End Property

        <DefaultValue(DefaultServicePassPhrase)> _
    <Category("Service")> _
    <Description("PassPhrase for connecting to VSP Service")> _
    <DisplayName("ServicePassPhrase")> _
    Public Property ServicePassPhrase() As String
            Get
                Return Me._servicePassPhrase
            End Get
            Set(ByVal value As String)
                Me._servicePassPhrase = value
            End Set
        End Property
    End Class