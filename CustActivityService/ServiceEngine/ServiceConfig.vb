Imports System.ComponentModel
Imports System.Xml

Public Class ServiceConfig

    ' IMPORTANT - remember to change the following of ServiceStartup project
    ' 1) Assembly Name (in Properties)
    ' 2) ConfigFileName in ServiceStartup.vb
    ' 3) ServiceName, DisplayName and Description in ServiceInstallation.vb

    Private Const DefaultTNS As String = "TST"

    Private sTNS As String = DefaultTNS
    Private sUID As String = DefaultTNS
    Private sPWD As String = DefaultTNS

    Private sFileFolder As String = String.Empty
    Private sStartEmails As String = "1400"
    Private sEmailDay As String = "SUN"
    Private sCCEmail As String = String.Empty
    Private sLastTimeExecuted As String = String.Empty

    Private _driveLetter As String = String.Empty
    Private _driveLetterIP As String = String.Empty


    Public Sub New()
        'Get settings from folder

        sFileFolder = My.Application.Info.DirectoryPath
        If sFileFolder.EndsWith("\") Then sFileFolder &= "\"

        If My.Computer.FileSystem.FileExists(My.Application.Info.DirectoryPath & "\SvcConfig.xml") Then
            Using xReader As Xml.XmlTextReader = New XmlTextReader(My.Application.Info.DirectoryPath & "\SvcConfig.xml")
                Do While xReader.Read()
                    Select Case xReader.NodeType
                        Case XmlNodeType.Element
                            Select Case xReader.Name
                                Case "DefaultTNS"
                                    sTNS = xReader.ReadElementContentAsString()
                                Case "DefaultUID"
                                    sUID = xReader.ReadElementContentAsString()
                                Case "DefaultPWD"
                                    sPWD = xReader.ReadElementContentAsString()
                                Case "StartEmail"
                                    sStartEmails = xReader.ReadElementContentAsString()
                                Case "EmailDay"
                                    sEmailDay = xReader.ReadElementContentAsString()
                                Case "CCEmail"
                                    sCCEmail = xReader.ReadElementContentAsString()
                                Case "LastTimeExecuted"
                                    sLastTimeExecuted = xReader.ReadElementContentAsString()
                                Case "DriveLetter"
                                    _driveLetter = xReader.ReadElementContentAsString()
                                Case "DriveLetterIP"
                                    _driveLetterIP = xReader.ReadElementContentAsString()
                            End Select
                    End Select
                Loop
                xReader.Close()
            End Using
        End If

        If sLastTimeExecuted.Length = 0 Then
            sLastTimeExecuted = DateAdd(DateInterval.Year, -1, DateTime.Now).ToString
        End If

    End Sub

    <DefaultValue(DefaultTNS)> _
    <Category("Oracle")> _
    <Description("This is the TNS Name of the Oracle Database that the Service write results to")> _
    <DisplayName("TNS Name of Oracle Schema to Log Results to")> _
    Public Property TNS() As String
        Get
            Return Me.sTNS
        End Get
        Set(ByVal value As String)
            Me.sTNS = value
        End Set
    End Property

    <DefaultValue(DefaultTNS)> _
    <Category("Oracle")> _
    <Description("This is the UID of the Oracle Database that the Service write results to")> _
    <DisplayName("UID of Oracle Schema to Log Results to")> _
    Public Property UID() As String
        Get
            Return Me.sUID
        End Get
        Set(ByVal value As String)
            Me.sUID = value
        End Set
    End Property

    <DefaultValue(DefaultTNS)> _
    <Category("Oracle")> _
    <Description("This is the PWD of the Oracle Schema UID")> _
    <DisplayName("PWD of Oracle Schema UID")> _
    Public Property PWD() As String
        Get
            Return Me.sPWD
        End Get
        Set(ByVal value As String)
            Me.sPWD = value
        End Set
    End Property

    <DefaultValue("1400")> _
    <Category("EmailStartTime")> _
    <Description("This is the time (military) to start sending emails")> _
    <DisplayName("Time to start emailing customers who have not ordered in the last week to sales reps")> _
    Public Property EmailStartTime() As String
        Get
            Return Me.sStartEmails
        End Get
        Set(ByVal value As String)
            Me.sStartEmails = value
        End Set
    End Property

    <DefaultValue("Application Directory")> _
    <Category("EmailStartTime")> _
    <Description("This is the Name of the FileFolder where Logs will be stored")> _
    <DisplayName("Name of the FileFolder where Logs will be stored")> _
    Public Property FileFolder() As String
        Get
            Return Me.sFileFolder
        End Get
        Set(ByVal value As String)
            Me.sFileFolder = value
        End Set
    End Property

    <DefaultValue("ALL")> _
    <Category("EmailStartTime")> _
    <Description("This is the Day to email sales reps")> _
    <DisplayName("Day to email stats to sales reps. ALL for everyday, else first three chars of day")> _
    Public Property EmailDay() As String
        Get
            Return sEmailDay
        End Get
        Set(ByVal value As String)
            sEmailDay = (value & String.Empty).ToString.ToUpper.Trim
        End Set
    End Property

    Public Property CCEmail() As String
        Get
            Return sCCEmail
        End Get
        Set(ByVal value As String)
            sCCEmail = value
        End Set
    End Property

    Public Property LastTimeExecuted() As String
        Get
            Return (sLastTimeExecuted)
        End Get
        Set(ByVal value As String)
            sLastTimeExecuted = value
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

    Public Function UpdateConfigNode(ByVal NodeName As String, ByVal NodeValue As String) As Boolean

        Try
            If My.Computer.FileSystem.FileExists(My.Application.Info.DirectoryPath & "\SvcConfig.xml") Then
                Dim MyXML As New XmlDocument()
                MyXML.Load(My.Application.Info.DirectoryPath & "\SvcConfig.xml")
                Dim MyXMLNode As XmlNode = MyXML.SelectSingleNode("/SvcConfig/" & NodeName)
                If MyXMLNode IsNot Nothing Then
                    MyXMLNode.ChildNodes(0).InnerText = NodeValue
                Else
                    MyXMLNode = MyXML.SelectSingleNode("/SvcConfig")
                    Dim elem As XmlNode = MyXML.CreateNode(XmlNodeType.Element, NodeName, Nothing)
                    elem.InnerText = NodeValue
                    MyXMLNode.AppendChild(elem)
                End If
                MyXML.Save(My.Application.Info.DirectoryPath & "\SvcConfig.xml")
                Return True
            End If

        Catch ex As Exception
            Return False
        End Try

        Return False

    End Function

End Class