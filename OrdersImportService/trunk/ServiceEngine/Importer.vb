Imports ServiceEngine.Extensions
Imports System.IO
Imports System.Xml
Imports System.Xml.Serialization
Imports System.Xml.Schema

Namespace OrdersImport

    Public Class SalesOrderImporter

        Private WithEvents importTimer As System.Threading.Timer
        Declare Function ProcessIdToSessionId Lib "kernel32.dll" (ByVal dwProcessId As Int32, ByRef pSessionId As Int32) As Int32

#Region "Service Variables"

        Private baseClass As ABSolution.ASFBASE1
        Private pricingClass As ABSolution.TACMAIN1
        Private DpdDefaultShipViaCode As String = String.Empty

        Private SOTINVH2_PC As String = String.Empty
        Private SOTORDR2_pricing As String = String.Empty
        Private SOTORDRP As String = String.Empty
        Private STAX_CODE_states As List(Of String)
        Private clsSOCORDR1 As TAC.SOCORDR1
        Private importInProcess As Boolean = False

        Private logFilename As String = String.Empty
        Private filefolder As String = String.Empty
        Private logStreamWriter As System.IO.StreamWriter

        Private DriveLetter As String = String.Empty
        Private DriveLetterIP As String = String.Empty
        Private convert As Boolean = False

        Private rowARTCUST1 As DataRow = Nothing
        Private rowARTCUST2 As DataRow = Nothing
        Private rowARTCUST3 As DataRow = Nothing
        Private rowWBTPARM1 As DataRow = Nothing

        Private rowSOTSVIA1 As DataRow = Nothing
        Private rowTATTERM1 As DataRow = Nothing
        Private rowICTITEM1 As DataRow = Nothing

        Private dst As DataSet
        Private vwXmlDataset As DataSet

        Private Const testMode As Boolean = False
        Private ImportErrorNotification As Hashtable
        Private DE_PARM_FOG_FREE_COST As Double = 0

        ' Header Errors
        Private Const ShipToOrderBlocked = "A"
        Private Const InvalidSoldTo = "B"
        Private Const ShipToClosed = "C"
        Private Const InvalidDPD = "D"
        Private Const InvalidDPDAddress = "E"
        Private Const InvalidSalesTax = "G"
        Private Const InvalidPricing = "H"
        Private Const InvalidSalesOrderTotal = "J"
        Private Const RequiresShipTo = "K"
        Private Const DpdCODCustomer = "P"
        Private Const ReviewSpecInstr = "R"
        Private Const InvalidShipTo = "S"
        Private Const InvalidTermsCode = "T"
        Private Const InvalidShipVia = "V"
        Private Const InvalidTaxCode = "X"
        Private Const DPDNoAnnualSupply = "9"

        'Detail Errors
        Private Const FrozenInactiveItem = "F"
        Private Const InvalidItem = "I"
        Private Const RevenueItemNoPrice = "N"
        Private Const QtyOrdered = "Q"
        Private Const InvalidUOM = "U"
        Private Const ItemAuthorizationError = "0"

        Private timerPeriod As Integer = 10

        Private ImportedFiles As List(Of String) = New List(Of String)
        Private ftpFileList As List(Of String) = New List(Of String)

        Private WithEvents Ftp1 As New nsoftware.IPWorks.Ftp

        Private SO_PARM_SHIP_ND As String = String.Empty
        Private SO_PARM_SHIP_ND_DPD As String = String.Empty
        Private SO_PARM_SHIP_ND_COD As String = String.Empty
        Private aspPriceCatgys As List(Of String)

        ' nSoftware License Keys
        Private nSoftwareZipkey As String = "315A4E384141315355425241533154453345383933333331580000000000000000000000000000003532323931555536000042424D454B375544463730460000"
        Private nSoftwareftpkey As String = "31504E384141315355425241533154453345383933333331580000000000000000000000000000003532323931555536000046384445474A3131583750500000"
        Private nSoftwareipportkey As String = "31504E38414131535542524153315445334538393333333158000000000000000000000000000000423252383654315A0000343854353048333650384B370000"
        Private nSoftwarepopkey As String = "31504E38414131535542524153315445334538393333333158000000000000000000000000000000394A37383437343200004E55523837363650504A42350000"

        Private Enum VisionWebStatus
            Received = 5
            OrderInProcess = 7
            WaitingForInformation = 10
            WaitingForFrame = 15
            RxLaunch = 20
            Surfacing = 25
            Treatment = 30
            Tinting = 35
            Finishing = 40
            Inspection = 45
            BreakageRedo = 50
            Shipping = 55
            Shipped = 60
            Cancelled = 900
            Other = 999
        End Enum

        Private Class Connection

            Public RemoteHost As String = String.Empty
            Public UserId As String = String.Empty
            Public Password As String = String.Empty

            Public LocalInDir As String = String.Empty
            Public LocalInDirArchive As String = String.Empty
            Public LocalOutDir As String = String.Empty
            Public LocalOutDirArchive As String = String.Empty

            Public RemoteInDirectory As String = String.Empty
            Public RemoteInDirectoryArchive As String = String.Empty
            Public RemoteOutDirectory As String = String.Empty
            Public RemoteOutDirectoryArchive As String = String.Empty

            Public Filename As String = String.Empty
            Public ConnectionDescription As String = String.Empty

            Public Sub New(ByVal OrderSource As String)
                Dim rowSOTPARMP As DataRow = ABSolution.ASCDATA1.GetDataRow("Select * From SOTPARMP WHERE SO_PARM_KEY = :PARM1", "V", New Object() {OrderSource})

                Dim svcConfig As New ServiceConfig
                Dim DriveLetter As String = svcConfig.DriveLetter.ToString.ToUpper.Trim
                Dim DriveLetterIP As String = svcConfig.DriveLetterIP.ToString.ToUpper.Trim
                Dim convert As Boolean = DriveLetter.Length > 0 AndAlso DriveLetterIP.Length > 0

                If rowSOTPARMP IsNot Nothing Then

                    LocalInDir = rowSOTPARMP.Item("SO_PARM_LOCAL_IN") & String.Empty
                    LocalInDir = LocalInDir.ToUpper.Trim
                    If convert And LocalInDir.StartsWith(DriveLetter) Then
                        LocalInDir = LocalInDir.Replace(DriveLetter, DriveLetterIP)
                    End If

                    If LocalInDir.Length > 0 AndAlso Not My.Computer.FileSystem.DirectoryExists(LocalInDir) Then
                        LocalInDir = String.Empty
                    End If
                    If LocalInDir.Length > 0 AndAlso Not LocalInDir.EndsWith("\") Then
                        LocalInDir &= "\"
                    End If

                    LocalInDirArchive = rowSOTPARMP.Item("SO_PARM_LOCAL_IN_ARCHIVE") & String.Empty
                    LocalInDirArchive = LocalInDirArchive.ToUpper.Trim
                    If convert And LocalInDirArchive.StartsWith(DriveLetter) Then
                        LocalInDirArchive = LocalInDirArchive.Replace(DriveLetter, DriveLetterIP)
                    End If

                    If LocalInDirArchive.Length > 0 AndAlso Not My.Computer.FileSystem.DirectoryExists(LocalInDirArchive) Then
                        LocalInDirArchive = String.Empty
                    End If
                    If LocalInDirArchive.Length > 0 AndAlso Not LocalInDirArchive.EndsWith("\") Then
                        LocalInDirArchive &= "\"
                    End If

                    LocalOutDir = rowSOTPARMP.Item("SO_PARM_LOCAL_OUT") & String.Empty
                    LocalOutDir = LocalOutDir.ToUpper.Trim
                    If convert And LocalOutDir.StartsWith(DriveLetter) Then
                        LocalOutDir = LocalOutDir.Replace(DriveLetter, DriveLetterIP)
                    End If

                    If LocalOutDir.Length > 0 AndAlso Not My.Computer.FileSystem.DirectoryExists(LocalOutDir) Then
                        LocalOutDir = String.Empty
                    End If
                    If LocalOutDir.Length > 0 AndAlso Not LocalOutDir.EndsWith("\") Then
                        LocalOutDir &= "\"
                    End If

                    LocalOutDirArchive = rowSOTPARMP.Item("SO_PARM_LOCAL_OUT_ARCHIVE") & String.Empty
                    LocalOutDirArchive = LocalOutDirArchive.ToUpper.Trim
                    If convert And LocalOutDirArchive.StartsWith(DriveLetter) Then
                        LocalOutDirArchive = LocalOutDirArchive.Replace(DriveLetter, DriveLetterIP)
                    End If

                    If LocalOutDirArchive.Length > 0 AndAlso Not My.Computer.FileSystem.DirectoryExists(LocalOutDirArchive) Then
                        LocalOutDirArchive = String.Empty
                    End If
                    If LocalOutDirArchive.Length > 0 AndAlso Not LocalOutDirArchive.EndsWith("\") Then
                        LocalOutDirArchive &= "\"
                    End If

                    RemoteInDirectory = rowSOTPARMP.Item("SO_PARM_REMOTE_IN") & String.Empty
                    RemoteInDirectoryArchive = rowSOTPARMP.Item("SO_PARM_REMOTE_IN_ARCHIVE") & String.Empty
                    RemoteOutDirectory = rowSOTPARMP.Item("SO_PARM_REMOTE_OUT") & String.Empty
                    RemoteOutDirectoryArchive = rowSOTPARMP.Item("SO_PARM_REMOTE_OUT_ARCHIVE") & String.Empty

                    RemoteHost = rowSOTPARMP.Item("SO_PARM_HOST") & String.Empty
                    UserId = rowSOTPARMP.Item("SO_PARM_USER_NAME") & String.Empty
                    Password = rowSOTPARMP.Item("SO_PARM_PASSWORD") & String.Empty

                    Filename = rowSOTPARMP.Item("SO_PARM_FILE_NAME") & String.Empty
                    ConnectionDescription = rowSOTPARMP.Item("SO_PARM_DESC") & String.Empty
                End If
            End Sub
        End Class

        Private prevent_blank_selection As Boolean = False

#End Region

#Region "Instaniate Service"

        Public Sub New()

        End Sub

#End Region

#Region "Data Management"

        Private Sub MainProcess()
            Try

                ' Prevent the code from firing if still importing
                If importInProcess Then Exit Sub

                If Not OpenLogFile() Then
                    Exit Sub
                End If

                ' Place a blank line in file to better see
                ' where each call starts.
                RecordLogEntry(String.Empty)
                RecordLogEntry("Enter MainProcess.")

                System.Threading.Thread.Sleep(2000)
                If LogIntoDatabase() Then
                    System.Threading.Thread.Sleep(2000)
                    If InitializeSettings() Then
                        System.Threading.Thread.Sleep(2000)
                        If PrepareDatasetEntries() Then
                            System.Threading.Thread.Sleep(2000)
                            If Not importInProcess Then
                                importInProcess = True
                                ProcessSalesOrders()
                                importInProcess = False
                            End If
                        End If
                    End If
                End If

                If testMode Then RecordLogEntry("Exit MainProcess.")
                RecordLogEntry("Closing Log file.")
                CloseLog()

            Catch ex As Exception
                importInProcess = False
                RecordLogEntry("MainProcess: " & ex.Message)
            Finally
                importInProcess = False
                If ABSolution.ASCMAIN1.oraCon.State = ConnectionState.Open Then
                    ABSolution.ASCMAIN1.oraCon.Close()
                End If
                DisposeOPD()
            End Try

        End Sub

        Public Sub LogIn()

            ' Start Service every 1 hours.
            ' This logic should have the service start on every hour. I added an extra 2 minutes
            Dim startInMinutes As Integer = ((60 - DateTime.Now.Minute) + 2) * 60000
            Dim startEvery As Integer = timerPeriod * 1000 * 60 ' every period Minutes

            If My.Application.Info.DirectoryPath.ToUpper.StartsWith("C:\VS") Then
                ' Give extra time when in test mode. usually stepping through code
                startEvery = 3 * 1000 * 60 * 60 ' 3 hours to avoid a restart when testing
                importTimer = New System.Threading.Timer _
                (New System.Threading.TimerCallback(AddressOf MainProcess), Nothing, 3000, startEvery)
            Else
                ' Orders Import should start right away.
                importTimer = New System.Threading.Timer _
                    (New System.Threading.TimerCallback(AddressOf MainProcess), Nothing, 3000, startEvery)
            End If

        End Sub

        Private Sub StartingProcess()
            ' Do nothing. just a way to start the service
        End Sub

        Private Function LogIntoDatabase() As Boolean
            LogIntoDatabase = False

            Try

                If testMode Then RecordLogEntry("Enter LogIntoDatabase.")

                Dim svcConfig As New ServiceConfig
                ABSolution.ASCMAIN1.DBS_COMPANY = svcConfig.UID
                ABSolution.ASCMAIN1.DBS_PASSWORD = svcConfig.PWD
                ABSolution.ASCMAIN1.DBS_SERVER = svcConfig.TNS

                timerPeriod = svcConfig.TimerPeriod
                If timerPeriod < 0 Then timerPeriod = 10
                If timerPeriod > 60 Then timerPeriod = 60

                If ABSolution.ASCMAIN1.DBS_COMPANY = "" OrElse ABSolution.ASCMAIN1.DBS_PASSWORD = "" OrElse ABSolution.ASCMAIN1.DBS_SERVER = "" Then
                    RecordLogEntry("LogIntoDatabase: Missing Credentials")
                    Return False
                End If

                If ABSolution.ASCMAIN1.oraCon.State = ConnectionState.Open Then
                    ABSolution.ASCMAIN1.oraCon.Close()
                End If

                Dim DEVELOPMENT_MACHINE_TNS As String = "(DESCRIPTION =(ADDRESS_LIST =(ADDRESS = (PROTOCOL = TCP)(HOST = localhost)(PORT = 1521)))(CONNECT_DATA = (SERVICE_NAME = orcl)))"
                DEVELOPMENT_MACHINE_TNS = ""

                If ABSolution.ASCMAIN1.DBS_TYPE = ABSolution.ASCMAIN1.DBS_TYPE_types.SQLServer Then
                    ABSolution.ASCMAIN1.oraCon.ConnectionString = "Data Source=" & IIf(ABSolution.ASCMAIN1.DBS_SERVER = "", ".", ABSolution.ASCMAIN1.DBS_SERVER) & ";Initial Catalog=" & ABSolution.ASCMAIN1.DBS_COMPANY & "; " & IIf(ABSolution.ASCMAIN1.DBS_SERVER = "", "User ID='" & ABSolution.ASCMAIN1.DBS_COMPANY & "'", "User ID='sa';Password='0ff1c3';") & ";Integrated Security=" & IIf(ABSolution.ASCMAIN1.DBS_SERVER = "", "True", "False") & ";MultipleActiveResultSets=True"
                Else
                    ABSolution.ASCMAIN1.oraCon.ConnectionString = "Data Source=" & IIf(ABSolution.ASCMAIN1.DBS_SERVER = "", DEVELOPMENT_MACHINE_TNS, ABSolution.ASCMAIN1.DBS_SERVER) & ";User ID=" & ABSolution.ASCMAIN1.DBS_COMPANY & ";Password=" & ABSolution.ASCMAIN1.DBS_PASSWORD & ";pooling=false"
                End If

                ABSolution.ASCMAIN1.oraCon.Open()
                ABSolution.ASCMAIN1.oraCmd = ABSolution.ASCMAIN1.oraCon.CreateCommand

                ABSolution.ASCMAIN1.oraSP.CommandType = CommandType.StoredProcedure
                ABSolution.ASCMAIN1.oraSP.Connection = ABSolution.ASCMAIN1.oraCon

                Dim myWorkstation As String = System.Net.Dns.GetHostName()
                Dim IPAddress As String = _
                System.Net.Dns.GetHostEntry(myWorkstation).AddressList(0).ToString()

                ABSolution.ASCMAIN1.DBS_IP_ADDRESS = IPAddress
                ABSolution.ASCMAIN1.DBS_SERVER_NAME = myWorkstation

                RecordLogEntry("Successful log into Oracle.")
                Return True

            Catch ex As Exception
                RecordLogEntry("LogIntoDatabase: " & ex.Message)
                Return False
            End Try

        End Function

        Private Function InitializeSettings() As Boolean

            Try

                Dim INIT_DATE As Date = DateTime.Now + ABSolution.ASCMAIN1.NowTSD

                If testMode Then RecordLogEntry("Enter InitializeSettings.")

                baseClass = New ABSolution.ASFBASE1
                pricingClass = New ABSolution.TACMAIN1

                DpdDefaultShipViaCode = String.Empty

                SOTINVH2_PC = String.Empty
                SOTORDR2_pricing = String.Empty
                SOTORDRP = String.Empty
                STAX_CODE_states = New List(Of String)

                logFilename = String.Empty
                filefolder = String.Empty

                rowARTCUST1 = Nothing
                rowARTCUST2 = Nothing
                rowARTCUST3 = Nothing

                rowSOTSVIA1 = Nothing
                rowTATTERM1 = Nothing
                rowICTITEM1 = Nothing

                dst = New DataSet

                ABSolution.ASCMAIN1.USER_ID = "service"

                ABSolution.ASCMAIN1.Set_DBS_Dependent_Strings()

                ABSolution.ASCMAIN1.SESSION_NO = ABSolution.ASCMAIN1.Next_Control_No("ASTLOGS1.SESSION_NO")
                If ABSolution.ASCMAIN1.DBS_TYPE = ABSolution.ASCMAIN1.DBS_TYPE_types.SQLServer Then
                    ABSolution.ASCMAIN1.DBS_SESSION_ID = 1
                Else
                    Dim rowSession As DataRow = ABSolution.ASCDATA1.GetDataRow("Select UserEnv('SESSIONID'), UserEnv('TERMINAL') from DUAL")
                    ABSolution.ASCMAIN1.DBS_SESSION_ID = rowSession.Item(0)
                End If
                ABSolution.ASCMAIN1.COMPUTER_NAME = My.Computer.Name


                ABSolution.ASCMAIN1.Get_Current_YP()

                ABSolution.ASCMAIN1.sql = "Select * from ASTPARM1 where AS_PARM_KEY = 'Z'"
                Dim tblASTPARM1 As DataTable = ABSolution.ASCDATA1.GetDataTable
                ABSolution.ASCMAIN1.rowASTPARM1 = tblASTPARM1.Rows(0)
                ABSolution.ASCMAIN1.tblASTFFMT1 = ABSolution.ASCDATA1.GetDataTable("*", "ASTFFMT1")
                ABSolution.ASCMAIN1.Temp_Table_Cleanup()

                Dim tblASTOPST1 As New DataTable
                With ABSolution.ASCDATA1.GetDataAdapter(tblASTOPST1, "ASTOPST1", "*", True, -1, False)
                    Dim rowASTOPST1 As DataRow = tblASTOPST1.NewRow
                    rowASTOPST1.Item("USER_ID") = ABSolution.ASCMAIN1.USER_ID
                    rowASTOPST1.Item("SESSION_NO") = ABSolution.ASCMAIN1.SESSION_NO
                    rowASTOPST1.Item("INIT_DATE") = INIT_DATE
                    rowASTOPST1.Item("YYYYPP") = ABSolution.ASCMAIN1.CYP
                    rowASTOPST1.Item("SELECTION_NO") = 0
                    rowASTOPST1.Item("RE_XNO") = 0
                    rowASTOPST1.Item("PRD_CLOSE_IND") = ABSolution.ASCMAIN1.EOM
                    rowASTOPST1.Item("FORM_INSTANCE_NO") = ABSolution.ASCMAIN1.Next_Control_No("ASFLOGON.FORM_INSTANCE_NO")
                    tblASTOPST1.Rows.Add(rowASTOPST1)
                    .Update(tblASTOPST1)
                    .Dispose()
                End With

                Dim tblASTLOGS1 As New DataTable
                With ABSolution.ASCDATA1.GetDataAdapter(tblASTLOGS1, "ASTLOGS1", "*", True, -1, False)
                    Dim rowASTLOGS1 As DataRow = tblASTLOGS1.NewRow
                    rowASTLOGS1.Item("SESSION_NO") = ABSolution.ASCMAIN1.SESSION_NO
                    rowASTLOGS1.Item("USER_ID") = ABSolution.ASCMAIN1.USER_ID
                    rowASTLOGS1.Item("SESSION_ID") = ABSolution.ASCMAIN1.DBS_SESSION_ID
                    rowASTLOGS1.Item("COMPUTER_NAME") = ABSolution.ASCMAIN1.COMPUTER_NAME
                    rowASTLOGS1.Item("DATE_LOGGED_ON") = INIT_DATE
                    rowASTLOGS1.Item("SESSION_STATUS") = "A"
                    tblASTLOGS1.Rows.Add(rowASTLOGS1)
                    .Update(tblASTLOGS1)
                    .Dispose()
                End With

                ' WTS Session ID
                ABSolution.ASCMAIN1.WTS_SESSION_ID = GetSessionId()
                If testMode Then RecordLogEntry("Exit InitializeSettings.")

                Return True

            Catch ex As Exception
                RecordLogEntry("InitializeSettings: " & ex.Message)
                Return False
            End Try

        End Function

        Public Function GetSessionId() As Int32
            Try
                Dim _currentProcess As Process = Process.GetCurrentProcess()
                Dim _processID As Int32 = _currentProcess.Id
                Dim _sessionID As Int32
                Dim _result As Boolean = ProcessIdToSessionId(_processID, _sessionID)
                Return _sessionID

            Catch ex As Exception
                Return 0
            End Try
        End Function

        Private Sub Ftp1_OnDirList(ByVal sender As System.Object, ByVal e As nsoftware.IPWorks.FtpDirListEventArgs) Handles Ftp1.OnDirList
            ftpFileList.Add(e.FileName)
        End Sub

        Private Sub ProcessSalesOrders()

            Try

                If testMode Then RecordLogEntry("Enter ProcessSalesOrders.")

                Dim svcConfig As New ServiceConfig
                DriveLetter = svcConfig.DriveLetter.ToString.ToUpper
                DriveLetterIP = svcConfig.DriveLetterIP.ToString.ToUpper
                convert = DriveLetter.Length > 0 AndAlso DriveLetterIP.Length > 0

                Dim ORDR_SOURCE As String = String.Empty

                ABSolution.ASCMAIN1.SESSION_NO = ABSolution.ASCMAIN1.Next_Control_No("ASTLOGS1.SESSION_NO", 1)

                If ABSolution.ASCMAIN1.ActiveForm Is Nothing Then
                    ABSolution.ASCMAIN1.ActiveForm = New ABSolution.ASFBASE1
                End If
                ABSolution.ASCMAIN1.ActiveForm.SELECTION_NO = "1"

                ' Initialize so we get a fresh copy for each processing cycle
                aspPriceCatgys = New List(Of String)
                rowWBTPARM1 = Nothing

                For Each rowSOTPARMP As DataRow In ABSolution.ASCDATA1.GetDataTable("SELECT DISTINCT SO_PARM_KEY ORDR_SOURCE FROM SOTPARMP WHERE NVL(SO_PARM_USE_SERVICE, '0') = '1' ORDER BY SO_PARM_KEY").Select("", "ORDR_SOURCE")

                    ' Use inner Try/Catch so if one Import fails the others still execute
                    Try
                        ORDR_SOURCE = rowSOTPARMP.Item("ORDR_SOURCE") & String.Empty

                        ' House Keeping, common things to do
                        Ftp1 = New nsoftware.IPWorks.Ftp
                        System.Threading.Thread.Sleep(1000)
                        Ftp1.RuntimeLicense = nSoftwareftpkey

                        ftpFileList.Clear()
                        ImportedFiles.Clear()
                        baseClass.clsASCBASE1.Fill_Records("SOTSVIAF", ORDR_SOURCE)

                        If Not ABSolution.ASCMAIN1.Logical_Lock("IMPSVC01", ORDR_SOURCE, False, False, True, 1) Then
                            RecordLogEntry("Order Import Type: " & ORDR_SOURCE & " locked by previous instance.")
                            Continue For
                        End If

                        If Not ClearDataSetTables(True) Then
                            Continue For
                        End If

                        ImportErrorNotification = New Hashtable

                        ' *****************************************************************************
                        ' Make sure there is an Entry in ASTNOTE1 for each Ordr Source
                        ' *****************************************************************************
                        Select Case ORDR_SOURCE
                            Case "D"
                                ' Hard D to get separate parameters for Vision Web Digital Eyelab orders
                                ProcessVisionWebDELOrders("V")

                            Case "Y"  ' (Y) Eyeconic
                                ProcessWebServiceSalesOrders(ORDR_SOURCE)

                            Case "X" ' AnyLens (A), DBVISION (D), 
                                '   do not grab Y for eyeconic. There order source is 'Y' it is a differnt import
                                '   do not grab O for OptiPort. There order source is 'O' it is a differnt import
                                For Each row As DataRow In ABSolution.ASCDATA1.GetDataTable("SELECT ORDR_LINE_SOURCE FROM XMTXREF1 WHERE ORDR_SOURCE = 'X' AND ORDR_LINE_SOURCE NOT IN ('Y', 'O')").Rows
                                    ProcessWebServiceSalesOrders(row.Item("ORDR_LINE_SOURCE") & String.Empty)

                                    If ImportErrorNotification.Keys.Count > 0 Then
                                        For Each Item As DictionaryEntry In ImportErrorNotification
                                            emailErrors(Item.Key, Item.Value)
                                        Next
                                    End If

                                    ImportErrorNotification.Clear()
                                Next

                            Case "U", "F" ' (U) AcuityLogic, (F) eyeFinity - Replaces SOFORDRF
                                ProcessEyeFinAquitySalesOrders(ORDR_SOURCE)

                            Case "E" 'EDI - (S) Spectera
                                ProcessEDISalesOrders(ORDR_SOURCE)

                            Case "B" ' US Vision Care Scan
                                ProcessBLScan(ORDR_SOURCE)
                                ' The sales order has an E as the order source
                                If ImportErrorNotification.Count > 0 Then
                                    ImportErrorNotification.Add("B", ImportErrorNotification.Item("E"))
                                    ImportErrorNotification.Remove("E")
                                End If

                            Case "V" 'Vision Web
                                ProcessVisionWebSalesOrders(ORDR_SOURCE)
                                System.Threading.Thread.Sleep(2000)
                                ExportVisionWebStatus()

                            Case "O" ' Optiport
                                ProcessOptiportSalesOrders(ORDR_SOURCE)
                        End Select

                    Catch ex As Exception
                        RecordLogEntry("ProcessSalesOrders, Order Source: " & ORDR_SOURCE & " - " & ex.Message)

                    Finally
                        ABSolution.ASCMAIN1.MultiTask_Release(, , 1)
                        Ftp1.Dispose()

                        If ImportErrorNotification.Keys.Count > 0 Then
                            For Each Item As DictionaryEntry In ImportErrorNotification
                                emailErrors(Item.Key, Item.Value)
                            Next
                        End If

                        ImportErrorNotification.Clear()
                    End Try
                Next

                If testMode Then RecordLogEntry("Exit ProcessSalesOrders.")

            Catch ex As Exception
                RecordLogEntry("ProcessSalesOrders: " & ex.Message)
            Finally
                ' Clean up any locks if an error occurs
                ABSolution.ASCMAIN1.MultiTask_Release(, , 1)
            End Try

        End Sub

        Private Sub ProcessWebServiceSalesOrders(ByVal sqlOrdrLineSource As String)

            Dim rowSOTORDRX As DataRow = Nothing
            Dim salesOrdersProcessed As Int16 = 0

            Dim ORDR_NO As String = String.Empty
            Dim ORDR_LNO As Int16 = 0

            Dim XML_ORDR_SOURCE As String = String.Empty
            Dim ORDR_LINE_SOURCE As String = String.Empty
            Dim CREATE_SHIP_TO As Boolean = False
            Dim SELECT_SHIP_TO_BY_TELE As Boolean = False
            Dim CALLER_NAME As String = String.Empty
            Dim ORDR_SHIP_COMPLETE As String = String.Empty
            Dim sql As String = String.Empty
            Dim webConnection As New Connection(sqlOrdrLineSource)

            ' Always X in this module. The Order Line Source determines the customer
            Dim ORDR_SOURCE As String = "X"

            Try
                If testMode Then RecordLogEntry("Enter ProcessWebServiceSalesOrders.")

                sql = "SELECT * FROM XSTORDR1 WHERE NVL(PROCESS_IND, '0') = '0'"
                sql &= " AND ORDR_SOURCE = (SELECT XML_ORDR_SOURCE FROM XMTXREF1 WHERE ORDR_LINE_SOURCE = '" & sqlOrdrLineSource & "')"

                baseClass.clsASCBASE1.Fill_Records("XSTORDR1", String.Empty, True, sql)

                If dst.Tables("XSTORDR1").Rows.Count = 0 Then
                    Exit Sub
                End If

                For Each rowXSTORDR1 As DataRow In dst.Tables("XSTORDR1").Select("", "XS_DOC_SEQ_NO")
                    ClearDataSetTables(False)
                    ORDR_LNO = 0
                    XML_ORDR_SOURCE = rowXSTORDR1.Item("ORDR_SOURCE") & String.Empty

                    If dst.Tables("XMTXREF1").Select("XML_ORDR_SOURCE = '" & XML_ORDR_SOURCE & "'", "").Length = 0 Then
                        RecordLogEntry("ORDR_SOURCE not found in XMTXREF1 for XS_DOC_SEQ_NO: " & rowXSTORDR1.Item("XS_DOC_SEQ_NO"))
                        rowXSTORDR1.Item("PROCESS_IND") = "E"
                        emailErrors("X", 1, "Order Source not found in XMTXREF1 for order " & rowXSTORDR1.Item("XS_DOC_SEQ_NO"))
                        Continue For
                    End If

                    With dst.Tables("XMTXREF1").Select("XML_ORDR_SOURCE = '" & XML_ORDR_SOURCE & "'", "")(0)
                        ORDR_LINE_SOURCE = .Item("ORDR_LINE_SOURCE") & String.Empty
                        ORDR_SOURCE = .Item("ORDR_SOURCE") & String.Empty
                        CREATE_SHIP_TO = (.Item("CREATE_SHIP_TO") & String.Empty) = "1"
                        SELECT_SHIP_TO_BY_TELE = (.Item("SELECT_SHIP_TO_BY_TELE") & String.Empty) = "1"
                        CALLER_NAME = .Item("CALLER_NAME") & String.Empty
                        ORDR_SHIP_COMPLETE = .Item("ORDR_SHIP_COMPLETE") & String.Empty
                        If ORDR_SHIP_COMPLETE.Length = 0 Then ORDR_SHIP_COMPLETE = "0"
                    End With

                    ' Flag entry as getting processed
                    rowXSTORDR1.Item("PROCESS_IND") = "1"

                    baseClass.clsASCBASE1.Fill_Records("XSTORDR2", (rowXSTORDR1.Item("XS_DOC_SEQ_NO") & String.Empty).ToString.Trim)

                    ' No details so get the hell out of here. We should charge $50.00 for processing and handling
                    If dst.Tables("XSTORDR2").Rows.Count = 0 Then
                        RecordLogEntry("No Web Service Sales Orders Details for XS Doc Seq No: " & (rowXSTORDR1.Item("XS_DOC_SEQ_NO") & String.Empty).ToString.Trim)
                        Continue For
                    End If

                    For Each rowXSTORDR2 As DataRow In dst.Tables("XSTORDR2").Select("", "XS_DOC_SEQ_LNO")
                        rowSOTORDRX = dst.Tables("SOTORDRX").NewRow

                        rowSOTORDRX.Item("ORDR_NO") = rowXSTORDR1.Item("XS_DOC_SEQ_NO") & String.Empty
                        rowSOTORDRX.Item("CUST_CODE") = rowXSTORDR1.Item("CUSTOMER_ID") & String.Empty
                        rowSOTORDRX.Item("CUST_SHIP_TO_NO") = rowXSTORDR1.Item("OFFICE_ID") & String.Empty

                        If IsDate(rowXSTORDR1.Item("TRANSMIT_DATE") & String.Empty) Then
                            rowSOTORDRX.Item("ORDR_DATE") = rowXSTORDR1.Item("TRANSMIT_DATE") & String.Empty
                        Else
                            rowSOTORDRX.Item("ORDR_DATE") = DateTime.Now.ToString("MM/dd/yyyy")
                        End If

                        If rowXSTORDR1.Item("SHIP_TO_PATIENT") & String.Empty = "Y" Then
                            rowSOTORDRX.Item("ORDR_DPD") = "1"
                        Else
                            rowSOTORDRX.Item("ORDR_DPD") = "0"
                        End If
                        rowSOTORDRX.Item("SHIP_VIA_CODE") = rowXSTORDR1.Item("SHIPPING_METHOD") & String.Empty
                        rowSOTORDRX.Item("PATIENT_NAME") = TruncateField((rowXSTORDR2.Item("PATIENT_NAME") & String.Empty).ToString.Trim, "SOTORDR2", "PATIENT_NAME")

                        rowSOTORDRX.Item("EDI_CUST_REF_NO") = rowXSTORDR1.Item("ORDER_ID") & String.Empty
                        rowSOTORDRX.Item("ORDR_SOURCE") = ORDR_SOURCE
                        rowSOTORDRX.Item("ORDR_TYPE_CODE") = "REG"

                        ORDR_LNO += 1
                        rowSOTORDRX.Item("ORDR_LNO") = ORDR_LNO
                        rowXSTORDR2.Item("ORDR_LNO") = ORDR_LNO

                        rowSOTORDRX.Item("CUST_LINE_REF") = TruncateField(rowXSTORDR2.Item("ITEM_ID") & String.Empty, "SOTORDR2", "CUST_LINE_REF")
                        rowSOTORDRX.Item("ORDR_QTY") = Val(rowXSTORDR2.Item("ORDER_QTY") & String.Empty)

                        ' Do not place in sotordr2
                        rowSOTORDRX.Item("ORDR_UNIT_PRICE_PATIENT") = 0 'Val(rowXSTORDR2.Item("PATIENT_PRICE") & String.Empty)

                        Select Case rowXSTORDR2.Item("ITEM_EYE") & String.Empty
                            Case "OD"
                                rowSOTORDRX.Item("ORDR_LR") = "R"
                            Case "OS"
                                rowSOTORDRX.Item("ORDR_LR") = "L"
                        End Select

                        rowSOTORDRX.Item("ORDR_LINE_SOURCE") = ORDR_LINE_SOURCE
                        rowSOTORDRX.Item("PRICE_CATGY_CODE") = rowXSTORDR2.Item("PRODUCT_KEY") & String.Empty
                        rowSOTORDRX.Item("ITEM_PROD_ID") = rowXSTORDR2.Item("UPC_CODE") & String.Empty
                        'rowSOTORDRX.Item("ITEM_BASE_CURVE") = rowXSTORDR2.Item("ITEM_BASE_CURVE") & String.Empty
                        'rowSOTORDRX.Item("ITEM_SPHERE_POWER") = rowXSTORDR2.Item("ITEM_SPHERE_POWER") & String.Empty
                        'rowSOTORDRX.Item("ITEM_CYLINDER") = rowXSTORDR2.Item("ITEM_CYLINDER") & String.Empty
                        'rowSOTORDRX.Item("ITEM_AXIS") = rowXSTORDR2.Item("ITEM_AXIS") & String.Empty
                        'rowSOTORDRX.Item("ITEM_DIAMETER") = rowXSTORDR2.Item("ITEM_DIAMETER") & String.Empty
                        'rowSOTORDRX.Item("ITEM_ADD_POWER") = rowXSTORDR2.Item("ITEM_ADD_POWER") & String.Empty
                        'rowSOTORDRX.Item("ITEM_COLOR") = rowXSTORDR2.Item("ITEM_COLOR") & String.Empty

                        rowSOTORDRX.Item("CUST_SHIP_TO_NAME") = TruncateField(rowXSTORDR1.Item("OFFICE_NAME") & String.Empty, "ARTCUST2", "CUST_SHIP_TO_NAME")
                        rowSOTORDRX.Item("CUST_SHIP_TO_PHONE") = TruncateField(rowXSTORDR1.Item("OFFICE_PHONE") & String.Empty, "ARTCUST2", "CUST_SHIP_TO_PHONE")
                        rowSOTORDRX.Item("CUST_SHIP_TO_ADDR1") = TruncateField(rowXSTORDR1.Item("OFFICE_SHIP_TO_ADDRESS1") & String.Empty, "ARTCUST2", "CUST_SHIP_TO_ADDR1")
                        rowSOTORDRX.Item("CUST_SHIP_TO_ADDR2") = TruncateField(rowXSTORDR1.Item("OFFICE_SHIP_TO_ADDRESS2") & String.Empty, "ARTCUST2", "CUST_SHIP_TO_ADDR2")
                        rowSOTORDRX.Item("CUST_SHIP_TO_ADDR3") = String.Empty
                        rowSOTORDRX.Item("CUST_SHIP_TO_CITY") = TruncateField(rowXSTORDR1.Item("OFFICE_SHIP_TO_CITY") & String.Empty, "ARTCUST2", "CUST_SHIP_TO_CITY")
                        rowSOTORDRX.Item("CUST_SHIP_TO_STATE") = TruncateField(rowXSTORDR1.Item("OFFICE_SHIP_TO_STATE") & String.Empty, "ARTCUST2", "CUST_SHIP_TO_STATE")
                        rowSOTORDRX.Item("CUST_SHIP_TO_ZIP_CODE") = TruncateField(rowXSTORDR1.Item("OFFICE_SHIP_TO_ZIP") & String.Empty, "ARTCUST2", "CUST_SHIP_TO_ZIP_CODE")
                        rowSOTORDRX.Item("CUST_SHIP_TO_COUNTRY") = "US"

                        rowSOTORDRX.Item("BILLING_NAME") = rowXSTORDR1.Item("BILLING_NAME") & String.Empty
                        rowSOTORDRX.Item("BILLING_ADDRESS1") = rowXSTORDR1.Item("BILLING_ADDRESS1") & String.Empty
                        rowSOTORDRX.Item("BILLING_ADDRESS2") = rowXSTORDR1.Item("BILLING_ADDRESS2") & String.Empty
                        rowSOTORDRX.Item("BILLING_CITY") = rowXSTORDR1.Item("BILLING_CITY") & String.Empty
                        rowSOTORDRX.Item("BILLING_STATE") = rowXSTORDR1.Item("BILLING_STATE") & String.Empty
                        rowSOTORDRX.Item("BILLING_ZIP") = rowXSTORDR1.Item("BILLING_ZIP") & String.Empty

                        rowSOTORDRX.Item("CUST_NAME") = TruncateField(rowXSTORDR1.Item("SHIP_TO_NAME") & String.Empty, "SOTORDR5", "CUST_NAME")
                        rowSOTORDRX.Item("CUST_PHONE") = TruncateField(rowXSTORDR1.Item("SHIP_TO_PHONE") & String.Empty, "SOTORDR5", "CUST_PHONE")
                        rowSOTORDRX.Item("CUST_ADDR1") = TruncateField(rowXSTORDR1.Item("SHIP_TO_ADDRESS1") & String.Empty, "SOTORDR5", "CUST_ADDR1")
                        rowSOTORDRX.Item("CUST_ADDR2") = TruncateField(rowXSTORDR1.Item("SHIP_TO_ADDRESS2") & String.Empty, "SOTORDR5", "CUST_ADDR2")
                        rowSOTORDRX.Item("CUST_CITY") = TruncateField(rowXSTORDR1.Item("SHIP_TO_CITY") & String.Empty, "SOTORDR5", "CUST_CITY")
                        rowSOTORDRX.Item("CUST_STATE") = TruncateField(rowXSTORDR1.Item("SHIP_TO_STATE") & String.Empty, "SOTORDR5", "CUST_STATE")
                        rowSOTORDRX.Item("CUST_ZIP_CODE") = TruncateField(rowXSTORDR1.Item("SHIP_TO_ZIP") & String.Empty, "SOTORDR5", "CUST_ZIP_CODE")

                        ' Set Item Code, Item Desc and Item Desc2
                        rowSOTORDRX.Item("ITEM_CODE") = rowXSTORDR2.Item("ITEM_CODE") & String.Empty
                        rowICTITEM1 = baseClass.LookUp("ICTITEM1", rowXSTORDR2.Item("ITEM_CODE") & String.Empty)
                        If rowICTITEM1 Is Nothing Then
                            rowSOTORDRX.Item("ITEM_DESC") = String.Empty
                            rowSOTORDRX.Item("ITEM_DESC2") = String.Empty
                        Else
                            rowSOTORDRX.Item("ITEM_DESC") = rowICTITEM1.Item("ITEM_DESC") & String.Empty
                            rowSOTORDRX.Item("ITEM_DESC2") = rowICTITEM1.Item("ITEM_DESC2") & String.Empty
                        End If

                        ' As per Maria do not lock Ship Via
                        rowSOTORDRX.Item("ORDR_LOCK_SHIP_VIA") = "0"
                        rowSOTORDRX.Item("ORDR_COMMENT") = String.Empty
                        rowSOTORDRX.Item("OFFICE_WEBSITE") = rowXSTORDR1.Item("OFFICE_WEBSITE") & String.Empty

                        dst.Tables("SOTORDRX").Rows.Add(rowSOTORDRX)
                    Next

                    ORDR_NO = String.Empty
                    If CreateSalesOrder(ORDR_NO, CREATE_SHIP_TO, SELECT_SHIP_TO_BY_TELE, ORDR_LINE_SOURCE, ORDR_SOURCE, CALLER_NAME, False, ORDR_SOURCE) Then
                        ' See if the Order gets free samples
                        AddSampleItemsToOrders(dst.Tables("SOTORDR1").Rows(0), dst.Tables("SOTORDR2"))

                        ' Reset Ship Complete flag since it may be changed by customer setup
                        dst.Tables("SOTORDR1").Rows(0).Item("ORDR_SHIP_COMPLETE") = ORDR_SHIP_COMPLETE
                        rowXSTORDR1.Item("ORDR_NO") = ORDR_NO

                        For Each rowXSTORDR2 As DataRow In dst.Tables("XSTORDR2").Rows
                            rowXSTORDR2.Item("ORDR_NO") = ORDR_NO
                        Next

                        UpdateDataSetTables()
                        salesOrdersProcessed += 1
                    Else
                        rowXSTORDR1.Item("PROCESS_IND") = "E"
                        RecordLogEntry("Web Service Doc Seq No: " & (rowXSTORDR1.Item("XS_DOC_SEQ_NO") & String.Empty).ToString & " not imported")
                        emailErrors("X", 1, "Web Service Doc Seq No: " & (rowXSTORDR1.Item("XS_DOC_SEQ_NO") & String.Empty).ToString & " not imported")

                        With baseClass
                            Try
                                .BeginTrans()
                                .clsASCBASE1.Update_Record_TDA("XSTORDR1")
                                .CommitTrans()
                            Catch ex As Exception
                                .Rollback()
                                RecordLogEntry("UpdateDataSetTables  : " & ex.Message)
                            End Try
                        End With
                    End If
                Next

                If testMode Then RecordLogEntry("Exit ProcessWebServiceSalesOrders.")

            Catch ex As Exception
                RecordLogEntry("ProcessWebServiceSalesOrders: " & ex.Message)
                emailErrors(ORDR_SOURCE, 1, ex.Message)
            Finally
                RecordLogEntry(salesOrdersProcessed & " " & IIf(webConnection.ConnectionDescription.Length = 0, "Web Service (" & sqlOrdrLineSource & ")", webConnection.ConnectionDescription) & " Sales Orders to process.")
            End Try

        End Sub

        ''' <summary>
        ''' Replaces ABSolution SOTORDRF
        ''' </summary>
        ''' <param name="ORDR_SOURCE"></param>
        ''' <remarks></remarks>
        Private Sub ProcessEyeFinAquitySalesOrders(ByVal ORDR_SOURCE As String)

            Dim ftpConnection As New Connection(ORDR_SOURCE)
            Dim orderFileName As String = String.Empty
            Dim orderData As String = String.Empty
            Dim orderElementsX() As String
            Dim orderElements() As String
            Dim tempStr As String = String.Empty
            Dim rowSOTORDRX As DataRow = Nothing
            Dim rowSOTSVIAF As DataRow = Nothing

            Dim ORDR_NO As String = String.Empty
            Dim CUST_CODE As String = String.Empty
            Dim CUST_SHIP_TO_NO As String = String.Empty
            Dim ORDR_LR As String = String.Empty
            Dim custShip As String = String.Empty
            Dim salesOrdersProcessed As Integer = 0
            Dim ITEM_DESC2 As String = String.Empty

            Dim orderNumber As New Hashtable
            Dim fileNumber As Integer = 0

            If testMode Then
                RecordLogEntry("ProcessEyeFinAquitySalesOrders LocalInDir: " & ftpConnection.LocalInDir)
                RecordLogEntry("ProcessEyeFinAquitySalesOrders LocalInDirArchive: " & ftpConnection.LocalInDirArchive)
                RecordLogEntry("ProcessEyeFinAquitySalesOrders LocalOutDir: " & ftpConnection.LocalOutDir)
                RecordLogEntry("ProcessEyeFinAquitySalesOrders LocalOutDirArchive: " & ftpConnection.LocalOutDirArchive)
                RecordLogEntry("ProcessEyeFinAquitySalesOrders RemoteInDirectory: " & ftpConnection.RemoteInDirectory)
                RecordLogEntry("ProcessEyeFinAquitySalesOrders RemoteInDirectoryArchive: " & ftpConnection.RemoteInDirectoryArchive)
                RecordLogEntry("ProcessEyeFinAquitySalesOrders RemoteOutDirectory: " & ftpConnection.RemoteOutDirectory)
                RecordLogEntry("ProcessEyeFinAquitySalesOrders RemoteOutDirectoryArchive: " & ftpConnection.RemoteOutDirectoryArchive)
            End If

            ' Perform FTP Here
            ftpFileList.Clear()
            ImportedFiles.Clear()

            Try
                Ftp1.User = ftpConnection.UserId
                Ftp1.Password = ftpConnection.Password
                Ftp1.RemoteHost = ftpConnection.RemoteHost
                Ftp1.Logon()

                ftpFileList = New List(Of String)
                If Ftp1.Connected Then
                    If ftpConnection.RemoteInDirectory.Length > 0 Then
                        Ftp1.RemotePath = "/" & ftpConnection.RemoteInDirectory
                    End If
                    Ftp1.ListDirectory()
                End If

                For Each fileFtp As String In ftpFileList

                    If fileFtp.Length = 0 Then Continue For
                    If testMode Then RecordLogEntry("ProcessEyeFinAquitySalesOrders ftp: " & fileFtp)

                    If Not fileFtp.EndsWith(".snt") Then Continue For
                    If Not fileFtp.StartsWith(ftpConnection.Filename) Then Continue For

                    Dim filePrefix As String = DateTime.Now.ToString("yyyyMMddhhmmss") & "_"

                    If testMode Then RecordLogEntry("ProcessEyeFinAquitySalesOrders ftp Download: " & fileFtp)
                    Ftp1.RemoteFile = fileFtp
                    Ftp1.LocalFile = ftpConnection.LocalInDir & filePrefix & fileFtp
                    Ftp1.Download()
                    RecordLogEntry("ProcessEyeFinAquitySalesOrders ftp Delete: " & fileFtp)
                    Ftp1.DeleteFile(fileFtp)

                    fileFtp = fileFtp.Replace(".snt", ".csv")
                    Ftp1.RemoteFile = fileFtp
                    Ftp1.LocalFile = ftpConnection.LocalInDir & filePrefix & fileFtp
                    Ftp1.Download()
                    If testMode Then RecordLogEntry("ProcessEyeFinAquitySalesOrders ftp Delete: " & fileFtp)
                    Ftp1.DeleteFile(fileFtp)
                Next

            Catch ex As Exception
                RecordLogEntry("ProcessEyeFinAquitySalesOrders ftp: " & ex.Message)
                emailErrors(ORDR_SOURCE, 1, ex.Message)
            Finally
                Ftp1.Logoff()
                Ftp1.Dispose()
            End Try

            Try
                For Each orderFile As String In My.Computer.FileSystem.GetFiles(ftpConnection.LocalInDir, FileIO.SearchOption.SearchTopLevelOnly, "*" & ftpConnection.Filename & "*.csv")

                    orderFileName = My.Computer.FileSystem.GetName(orderFile)
                    RecordLogEntry("Importing file: " & orderFileName)
                    fileNumber += 1

                    Using orderReader As New StreamReader(orderFile)

                        While orderReader.Peek <> -1

                            orderData = orderReader.ReadLine()
                            orderElementsX = orderData.Split(Chr(34) & Chr(34))
                            tempStr = String.Empty

                            ' Need to weed out the ""
                            For iCtr As Integer = 0 To orderElementsX.Count - 1
                                If iCtr Mod 2 = 1 Then
                                    orderElementsX(iCtr) = orderElementsX(iCtr).Replace(",", ".")
                                End If
                                tempStr &= orderElementsX(iCtr)
                            Next
                            orderElements = tempStr.Split(",")

                            rowSOTORDRX = dst.Tables("SOTORDRX").NewRow

                            ORDR_NO = orderElements(0) & String.Empty

                            ' '' A file contained and Order No TR and Tr. This is a duplicate in Oracle
                            ' '' Each one had a LNO 1 and 2 causing a duplicate key error.
                            ' '' To change Order Number in each file to be unique.
                            If ORDR_NO.Length <= 4 Then
                                ORDR_NO = orderElements(0) & "_" & fileNumber.ToString.Trim
                                If Not orderNumber.ContainsKey(ORDR_NO) Then
                                    orderNumber.Add(ORDR_NO, ORDR_NO & "_" & orderNumber.Count.ToString.Trim)
                                End If
                                ORDR_NO = orderNumber.Item(ORDR_NO)
                            End If

                            rowSOTORDRX.Item("ORDR_SOURCE") = ORDR_SOURCE
                            rowSOTORDRX.Item("ORDR_NO") = ORDR_NO
                            rowSOTORDRX.Item("ORDR_LNO") = orderElements(1)
                            rowSOTORDRX.Item("EDI_CUST_REF_NO") = TruncateField(orderElements(0), "SOTORDR1", "EDI_CUST_REF_NO")
                            rowSOTORDRX.Item("ORDR_LINE_SOURCE") = ORDR_SOURCE
                            rowSOTORDRX.Item("ORDR_TYPE_CODE") = "REG"

                            custShip = (orderElements(5) & String.Empty).ToString.Trim
                            CUST_CODE = String.Empty
                            CUST_SHIP_TO_NO = String.Empty

                            If custShip.Contains("-") Then
                                CUST_CODE = custShip.Split("-")(0) & String.Empty
                                CUST_SHIP_TO_NO = custShip.Split("-")(1) & String.Empty
                            ElseIf custShip.Length > 6 Then
                                CUST_CODE = custShip.Substring(0, 6)
                                CUST_SHIP_TO_NO = custShip.Substring(6)
                                CUST_SHIP_TO_NO = CUST_SHIP_TO_NO.Trim
                                If CUST_SHIP_TO_NO.Length > 6 Then
                                    CUST_SHIP_TO_NO = CUST_SHIP_TO_NO.Substring(0, 6)
                                End If
                            Else
                                CUST_CODE = custShip
                                CUST_SHIP_TO_NO = String.Empty
                            End If

                            CUST_CODE = CUST_CODE.Trim
                            CUST_SHIP_TO_NO = CUST_SHIP_TO_NO.Trim

                            If CUST_SHIP_TO_NO = "000000" Then
                                CUST_SHIP_TO_NO = String.Empty
                            End If

                            CUST_CODE = ABSolution.ASCMAIN1.Format_Field(CUST_CODE, "CUST_CODE")
                            If CUST_SHIP_TO_NO.Length > 0 Then
                                CUST_SHIP_TO_NO = ABSolution.ASCMAIN1.Format_Field(CUST_SHIP_TO_NO, "CUST_SHIP_TO_NO")
                            End If

                            rowSOTORDRX.Item("CUST_CODE") = CUST_CODE
                            rowSOTORDRX.Item("CUST_SHIP_TO_NO") = CUST_SHIP_TO_NO

                            rowSOTORDRX.Item("CUST_NAME") = TruncateField(orderElements(6), "SOTORDR1", "CUST_NAME")
                            rowSOTORDRX.Item("ORDR_DPD") = IIf(orderElements(4) = "1", "1", "0")
                            rowSOTORDRX.Item("CUST_LINE_REF") = TruncateField(orderElements(16), "SOTORDR1", "CUST_LINE_REF")
                            rowSOTORDRX.Item("ORDR_QTY") = Val(orderElements(17) & String.Empty)
                            If orderElements(21).Trim.Length > 0 Then
                                rowSOTORDRX.Item("ITEM_CODE") = TruncateField(orderElements(21), "SOTORDR2", "ITEM_CODE")
                            Else
                                rowSOTORDRX.Item("ITEM_CODE") = TruncateField(orderElements(20), "SOTORDR2", "ITEM_CODE")
                            End If


                            ORDR_LR = (orderElements(46) & String.Empty).Trim.ToUpper
                            If ORDR_LR.Length > 2 Then ORDR_LR = ORDR_LR.Substring(0, 2)
                            Select Case ORDR_LR
                                Case "OD", "RI"
                                    rowSOTORDRX.Item("ORDR_LR") = "R"
                                Case "OS", "LE"
                                    rowSOTORDRX.Item("ORDR_LR") = "L"
                                Case "OU", "BO"
                                    rowSOTORDRX.Item("ORDR_LR") = "B"
                            End Select

                            rowSOTORDRX.Item("ITEM_DESC") = TruncateField(orderElements(19), "SOTORDR2", "ITEM_DESC")

                            ' Item Desc 2 contains attributes to assist selection of product if invalid item code
                            ITEM_DESC2 = "/" & orderElements(22) '  Base Curve
                            ITEM_DESC2 &= "/" & orderElements(24) ' Sphere
                            ITEM_DESC2 &= "/" & orderElements(26) ' Cylinder
                            ITEM_DESC2 &= "/" & orderElements(27) ' Axis
                            ITEM_DESC2 &= "/" & orderElements(23) ' Diameter
                            ITEM_DESC2 &= "/" & orderElements(25) ' Add Power
                            ITEM_DESC2 &= "/" & orderElements(28) ' Color
                            rowSOTORDRX.Item("ITEM_DESC2") = TruncateField(ITEM_DESC2, "SOTORDR2", "ITEM_DESC2")

                            rowSOTORDRX.Item("ORDR_DATE") = DateTime.Now.ToString("dd-MMM-yyyy")
                            rowSOTORDRX.Item("ORDR_CALLER_NAME") = StrConv(TruncateField(orderElements(7), "SOTORDR1", "ORDR_CALLER_NAME"), VbStrConv.ProperCase)
                            rowSOTORDRX.Item("ORDR_LOCK_SHIP_VIA") = "0"

                            ' See if we are to use a specified shipping method
                            If dst.Tables("SOTSVIAF").Select("SHIP_VIA_DESC = '" & orderElements(32) & "'", "").Length > 0 Then
                                rowSOTSVIAF = dst.Tables("SOTSVIAF").Select("SHIP_VIA_DESC = '" & orderElements(32) & "'", "")(0)

                                If rowSOTORDRX.Item("ORDR_DPD") = "1" Then
                                    rowSOTORDRX.Item("SHIP_VIA_CODE") = rowSOTSVIAF.Item("SHIP_VIA_CODE_DPD") & String.Empty
                                Else
                                    rowSOTORDRX.Item("SHIP_VIA_CODE") = rowSOTSVIAF.Item("SHIP_VIA_CODE") & String.Empty
                                End If

                                ' As per Maria, Lock Ship Via if they send use a Specific Ship Via. 20110926
                                If rowSOTORDRX.Item("SHIP_VIA_CODE") & String.Empty <> String.Empty Then
                                    rowSOTORDRX.Item("ORDR_LOCK_SHIP_VIA") = "1"
                                End If

                            ElseIf orderElements(32).Trim.ToUpper = "STANDARD" Then
                                ' Do nothing, the conversion to Sales Order will grab the ship via from the master tables.

                                'Dim rowARTCUST3 As DataRow = baseClass.LookUp("ARTCUST3", CUST_CODE)
                                'If rowARTCUST3 IsNot Nothing Then
                                '    If rowSOTORDRX.Item("ORDR_DPD") = "1" Then
                                '        rowSOTORDRX.Item("SHIP_VIA_CODE") = rowARTCUST3.Item("SHIP_VIA_CODE_DPD")
                                '    Else
                                '        rowSOTORDRX.Item("SHIP_VIA_CODE") = rowARTCUST3.Item("SHIP_VIA_CODE")
                                '    End If
                                'End If

                                '' If ship to the use the ship to Ship Via
                                'rowARTCUST2 = baseClass.LookUp("ARTCUST2", New String() {CUST_CODE, CUST_SHIP_TO_NO})

                                'If rowARTCUST2 IsNot Nothing AndAlso rowARTCUST2.Item("CUST_SHIP_TO_SHIP_VIA_CODE") & String.Empty <> String.Empty Then
                                '    If rowSOTORDRX.Item("ORDR_DPD") = "1" Then
                                '        If rowARTCUST2.Item("CUST_SHIP_TO_SHIP_VIA_CODE_DPD") & String.Empty <> String.Empty Then
                                '            rowSOTORDRX.Item("SHIP_VIA_CODE") = rowARTCUST2.Item("CUST_SHIP_TO_SHIP_VIA_CODE_DPD") & String.Empty
                                '        End If
                                '    Else
                                '        If rowARTCUST2.Item("CUST_SHIP_TO_SHIP_VIA_CODE") & String.Empty <> String.Empty Then
                                '            rowSOTORDRX.Item("SHIP_VIA_CODE") = rowARTCUST2.Item("CUST_SHIP_TO_SHIP_VIA_CODE") & String.Empty
                                '        End If
                                '    End If

                                'End If
                            End If

                            'If rowSOTORDRX.Item("SHIP_VIA_CODE") & String.Empty = String.Empty Then
                            '    rowSOTORDRX.Item("SHIP_VIA_CODE") = "SD"
                            'End If

                            ' Second part of Else - As per Maria if DPD the ship complete
                            If rowSOTORDRX.Item("ORDR_DPD") = "1" Then
                                rowSOTORDRX.Item("ORDR_SHIP_COMPLETE") = "1"
                            Else
                                rowSOTORDRX.Item("ORDR_SHIP_COMPLETE") = IIf(orderElements(33) = "1", "1", "0")
                            End If

                            rowSOTORDRX.Item("PATIENT_NAME") = orderElements(44)
                            If rowSOTORDRX.Item("ORDR_DPD") = "1" AndAlso (rowSOTORDRX.Item("PATIENT_NAME") & String.Empty).ToString.Length = 0 Then
                                rowSOTORDRX.Item("PATIENT_NAME") = orderElements(34)
                            End If

                            rowSOTORDRX.Item("PATIENT_NAME") = TruncateField(rowSOTORDRX.Item("PATIENT_NAME") & String.Empty, "SOTORDR2", "PATIENT_NAME")
                            rowSOTORDRX.Item("PATIENT_NAME") = StrConv(rowSOTORDRX.Item("PATIENT_NAME") & String.Empty, VbStrConv.ProperCase)

                            rowSOTORDRX.Item("ORDR_CUST_PO") = TruncateField(orderElements(45), "SOTORDR1", "ORDR_CUST_PO")
                            rowSOTORDRX.Item("CUST_PHONE") = TruncateField(orderElements(13), "SOTORDR5", "CUST_PHONE")

                            ' This is the Bill To
                            rowSOTORDRX.Item("CUST_SHIP_TO_NAME") = TruncateField(orderElements(34), "SOTORDR5", "CUST_NAME")
                            rowSOTORDRX.Item("CUST_SHIP_TO_ADDR1") = TruncateField(orderElements(35), "SOTORDR5", "CUST_ADDR1")
                            rowSOTORDRX.Item("CUST_SHIP_TO_ADDR2") = TruncateField(orderElements(36), "SOTORDR5", "CUST_ADDR2")
                            rowSOTORDRX.Item("CUST_SHIP_TO_CITY") = TruncateField(orderElements(37), "SOTORDR5", "CUST_CITY")
                            rowSOTORDRX.Item("CUST_SHIP_TO_STATE") = TruncateField(orderElements(38), "SOTORDR5", "CUST_STATE")
                            rowSOTORDRX.Item("CUST_SHIP_TO_ZIP_CODE") = TruncateField(orderElements(39), "SOTORDR5", "CUST_ZIP_CODE")
                            rowSOTORDRX.Item("CUST_SHIP_TO_PHONE") = TruncateField(orderElements(40), "SOTORDR5", "CUST_PHONE")
                            rowSOTORDRX.Item("CUST_SHIP_TO_FAX") = TruncateField(orderElements(41), "SOTORDR5", "CUST_FAX")
                            rowSOTORDRX.Item("CUST_SHIP_TO_EMAIL") = TruncateField(orderElements(42), "SOTORDR5", "CUST_EMAIL")
                            rowSOTORDRX.Item("CUST_SHIP_TO_COUNTRY") = "US"

                            'These fields are used to set the Ship To if the customer wants us to use the data they provided.
                            rowSOTORDRX.Item("CUST_NAME") = TruncateField(orderElements(34), "SOTORDR5", "CUST_NAME")
                            rowSOTORDRX.Item("CUST_ADDR1") = TruncateField(orderElements(35), "SOTORDR5", "CUST_ADDR1")
                            rowSOTORDRX.Item("CUST_ADDR2") = TruncateField(orderElements(36), "SOTORDR5", "CUST_ADDR2")
                            rowSOTORDRX.Item("CUST_CITY") = TruncateField(orderElements(37), "SOTORDR5", "CUST_CITY")
                            rowSOTORDRX.Item("CUST_STATE") = TruncateField(orderElements(38), "SOTORDR5", "CUST_STATE")
                            rowSOTORDRX.Item("CUST_ZIP_CODE") = TruncateField(orderElements(39), "SOTORDR5", "CUST_ZIP_CODE")
                            rowSOTORDRX.Item("CUST_PHONE") = TruncateField(orderElements(40), "SOTORDR5", "CUST_PHONE")
                            rowSOTORDRX.Item("CUST_FAX") = TruncateField(orderElements(41), "SOTORDR5", "CUST_FAX")
                            rowSOTORDRX.Item("CUST_EMAIL") = TruncateField(orderElements(42), "SOTORDR5", "CUST_EMAIL")
                            rowSOTORDRX.Item("CUST_COUNTRY") = "US"

                            ' If Aquity Logic and DPD and ShipTo Name is blank then use Patient name
                            If ORDR_SOURCE = "U" AndAlso rowSOTORDRX.Item("ORDR_DPD") = "1" AndAlso rowSOTORDRX.Item("CUST_NAME") & String.Empty = String.Empty Then
                                rowSOTORDRX.Item("CUST_NAME") = TruncateField(orderElements(44), "SOTORDR5", "CUST_NAME")
                            End If

                            ' Grab Customer Name from master tables when not a DPD, since we always use the 
                            ' Address provided as the Ship To Address
                            If rowSOTORDRX.Item("ORDR_DPD") & String.Empty <> "1" Then
                                rowARTCUST1 = baseClass.LookUp("ARTCUST1", CUST_CODE)
                                rowARTCUST2 = baseClass.LookUp("ARTCUST2", New String() {CUST_CODE, CUST_SHIP_TO_NO})

                                If CUST_SHIP_TO_NO.Length > 0 AndAlso rowARTCUST2 IsNot Nothing Then
                                    rowSOTORDRX.Item("CUST_NAME") = rowARTCUST2.Item("CUST_SHIP_TO_NAME") & String.Empty
                                ElseIf rowARTCUST1 IsNot Nothing Then
                                    rowSOTORDRX.Item("CUST_NAME") = rowARTCUST1.Item("CUST_NAME") & String.Empty
                                End If
                            End If

                            dst.Tables("SOTORDRX").Rows.Add(rowSOTORDRX)
                        End While
                    End Using

                    ImportedFiles.Add(orderFile)
                Next

                If dst.Tables("SOTORDRX").Rows.Count > 0 Then
                    ' Commit the data from the Excel file and then archive the file
                    Dim UpdateInProcess As Boolean = False
                    With baseClass
                        Try
                            .BeginTrans()
                            UpdateInProcess = True
                            .clsASCBASE1.Update_Record_TDA("SOTORDRX")
                            .CommitTrans()
                            UpdateInProcess = False

                        Catch ex As Exception
                            If UpdateInProcess Then .Rollback()
                            RecordLogEntry("ProcessEyeFinAquitySalesOrders: " & ex.Message)
                        End Try

                    End With

                End If

                Try
                    ' Move CSN and any SNT file extensions to the archive directory
                    For Each orderFile As String In ImportedFiles
                        If testMode Then RecordLogEntry("ProcessEyeFinAquitySalesOrders ftp Move file: " & orderFile)
                        My.Computer.FileSystem.MoveFile(orderFile, ftpConnection.LocalInDirArchive & My.Computer.FileSystem.GetName(orderFile))
                        orderFile = orderFile.Replace(".csv", ".snt")
                        If My.Computer.FileSystem.FileExists(orderFile) Then
                            If testMode Then RecordLogEntry("ProcessEyeFinAquitySalesOrders ftp Move file: " & orderFile)
                            My.Computer.FileSystem.MoveFile(orderFile, ftpConnection.LocalInDirArchive & My.Computer.FileSystem.GetName(orderFile))
                        End If
                    Next

                Catch ex As Exception
                    RecordLogEntry("ProcessEyeFinAquitySalesOrders: Move files to Archive" & ex.Message)
                End Try

                ' If DPD then set the LR
                ORDR_LR = String.Empty
                ORDR_NO = String.Empty
                Dim ORDR_CALLER_NAME As String = String.Empty
                Dim ORDR_SHIP_COMPLETE As String = String.Empty
                dst.Tables("SOTORDRX").Rows.Clear()

                ' Need to process each order individually for pricing reasons; therefore
                ' need to move the datat to a temp data table and process each order individually
                For Each headers As DataRow In ABSolution.ASCDATA1.GetDataTable("SELECT DISTINCT ORDR_NO FROM SOTORDRX WHERE PROCESS_IND IS NULL AND ORDR_SOURCE = :PARM1", String.Empty, "V", New Object() {ORDR_SOURCE}).Rows
                    ClearDataSetTables(True)
                    ORDR_NO = headers.Item("ORDR_NO") & String.Empty
                    baseClass.clsASCBASE1.Fill_Records("SOTORDRX", New Object() {ORDR_SOURCE, ORDR_NO})

                    If dst.Tables("SOTORDRX").Rows.Count = 0 Then
                        RecordLogEntry("ProcessEyeFinAquitySalesOrders: Invalid Order Number (" & ORDR_NO & ") for " & ftpConnection.ConnectionDescription)
                        Continue For
                    End If

                    If dst.Tables("SOTORDRX").Rows(0).Item("ORDR_DPD") & String.Empty = "1" _
                            AndAlso dst.Tables("SOTORDRX").Rows.Count = 2 _
                            AndAlso dst.Tables("SOTORDRX").Select("ISNULL(ORDR_LR, '*') <> '*'").Length = 0 Then
                        ' Since only 2 rows, first row is R second is L
                        ORDR_LR = "R"
                        For Each row As DataRow In dst.Tables("SOTORDRX").Select("", "ORDR_LNO")
                            row.Item("ORDR_LR") = ORDR_LR
                            ORDR_LR = "L"
                        Next
                    End If

                    ORDR_NO = String.Empty
                    ORDR_CALLER_NAME = dst.Tables("SOTORDRX").Rows(0).Item("ORDR_CALLER_NAME") & String.Empty
                    ORDR_SHIP_COMPLETE = dst.Tables("SOTORDRX").Rows(0).Item("ORDR_SHIP_COMPLETE") & String.Empty

                    ' For Aquity logic use address in ARTCUST2
                    Dim AlwaysUseImportedShipToAddress As Boolean = ORDR_SOURCE <> "U"

                    If CreateSalesOrder(ORDR_NO, False, False, ORDR_SOURCE, ORDR_SOURCE, ORDR_CALLER_NAME, AlwaysUseImportedShipToAddress, ORDR_SOURCE) Then
                        dst.Tables("SOTORDR1").Rows(0).Item("ORDR_SHIP_COMPLETE") = ORDR_SHIP_COMPLETE
                        UpdateDataSetTables()
                        salesOrdersProcessed += 1
                        Try
                            ABSolution.ASCDATA1.ExecuteSQL("DELETE FROM SOTORDRX WHERE ORDR_SOURCE = :PARM1 AND ORDR_NO = :PARM2", _
                                                           "VV", _
                                                        New Object() {ORDR_SOURCE, dst.Tables("SOTORDRX").Rows(0).Item("ORDR_NO") & String.Empty})

                        Catch ex As Exception

                        End Try
                    Else
                        RecordLogEntry("ProcessEyeFinAquitySalesOrders: " & "Could not create sales order for " & dst.Tables("SOTORDRX").Rows(0).Item("ORDR_NO"))
                    End If
                Next
            Catch ex As Exception
                RecordLogEntry("ProcessEyeFinAquitySalesOrders Loop Directory: " & ex.Message)
                emailErrors(ORDR_SOURCE, 1, ex.Message)
            Finally
                RecordLogEntry(salesOrdersProcessed & " " & ftpConnection.ConnectionDescription & " Sales Orders imported.")
            End Try
        End Sub

        ''' <summary>
        ''' Replaces ABSolution SOTORDRE
        ''' </summary>
        ''' <param name="ORDR_SOURCE"></param>
        ''' <remarks></remarks>
        Private Sub ProcessEDISalesOrders(ByVal ORDR_SOURCE As String)

            Dim sql As String = String.Empty

            Dim ediConnection As Connection = New Connection(ORDR_SOURCE)
            Dim rowSOTORDRX As DataRow = Nothing
            Dim rowWK As DataRow = Nothing
            Dim rowEDT850I0 As DataRow = Nothing

            Dim EDI_ISA_NO As String = String.Empty
            Dim EDI_BATCH_NO As String = String.Empty
            Dim EDI_DOC_SEQ_NO As String = String.Empty
            Dim ITEM_UPC_CODE As String = String.Empty

            Dim CUST_CODE As String = String.Empty
            Dim CUST_SHIP_TO_NO As String = String.Empty
            Dim PATIENT_NAME As String = String.Empty
            Dim rowEDT850I5 As DataRow = Nothing
            Dim strVal As String = String.Empty
            Dim salesOrdersProcessed As Integer = 0

            sql = "Select Distinct EDI_ISA_NO From EDT850I1 Where EDI_BATCH_NO IS NULL"
            baseClass.clsASCBASE1.Fill_Records("EDT850I1D", String.Empty, False, sql)

            Try
                For Each rowEDT850I1D As DataRow In dst.Tables("EDT850I1D").Select("", "EDI_ISA_NO")

                    EDI_ISA_NO = rowEDT850I1D.Item("EDI_ISA_NO") & String.Empty

                    sql = "Select * From EDT850I1 Where EDI_ISA_NO = '" & EDI_ISA_NO & "' AND EDI_BATCH_NO IS NULL"
                    baseClass.clsASCBASE1.Fill_Records("EDT850I1", String.Empty, False, sql)

                    EDI_BATCH_NO = ABSolution.ASCMAIN1.Next_Control_No("EDT850I0.EDI_BATCH_NO", 1)

                    rowWK = dst.Tables("EDT850I1").Select("EDI_ISA_NO = '" & EDI_ISA_NO & "'", "")(0)

                    rowEDT850I0 = dst.Tables("EDT850I0").NewRow
                    rowEDT850I0.Item("EDI_BATCH_NO") = EDI_BATCH_NO
                    rowEDT850I0.Item("EDI_BATCH_DATE_TIME") = DateTime.Now
                    rowEDT850I0.Item("EDI_TP_QUAL") = rowWK.Item("EDI_TP_QUAL") & String.Empty
                    rowEDT850I0.Item("EDI_TP_ID") = rowWK.Item("EDI_TP_ID") & String.Empty
                    rowEDT850I0.Item("EDI_OUR_ID") = rowWK.Item("EDI_OUR_ID") & String.Empty
                    rowEDT850I0.Item("EDI_OUR_QUAL") = rowWK.Item("EDI_OUR_QUAL") & String.Empty
                    rowEDT850I0.Item("EDI_PROCESS_IND") = String.Empty
                    rowEDT850I0.Item("EDI_BATCH_PROC_DATE_TIME") = DateTime.Now
                    rowEDT850I0.Item("EDI_BATCH_PROC_OPER") = ABSolution.ASCMAIN1.USER_ID
                    rowEDT850I0.Item("NO_OF_DOCS") = dst.Tables("EDT850I1").Rows.Count
                    dst.Tables("EDT850I0").Rows.Add(rowEDT850I0)

                    ' Header Data
                    For Each rowEDT850I1 As DataRow In dst.Tables("EDT850I1").Select("EDI_ISA_NO = '" & EDI_ISA_NO & "'", "EDI_DOC_SEQ_NO")

                        EDI_DOC_SEQ_NO = rowEDT850I1.Item("EDI_DOC_SEQ_NO") & String.Empty
                        CUST_CODE = String.Empty
                        CUST_SHIP_TO_NO = String.Empty

                        rowEDT850I1.Item("EDI_BATCH_NO") = EDI_BATCH_NO

                        baseClass.clsASCBASE1.Fill_Records("EDT850I2", EDI_DOC_SEQ_NO)
                        baseClass.clsASCBASE1.Fill_Records("EDT850I5", EDI_DOC_SEQ_NO)

                        ' Details Data
                        For Each rowEDT850I2 As DataRow In dst.Tables("EDT850I2").Select("", "EDI_DOC_SEQ_NO, EDI_DTL_SEQ")

                            rowSOTORDRX = dst.Tables("SOTORDRX").NewRow
                            rowSOTORDRX.Item("ORDR_SOURCE") = ORDR_SOURCE
                            rowSOTORDRX.Item("ORDR_NO") = EDI_DOC_SEQ_NO
                            rowSOTORDRX.Item("ORDR_LNO") = rowEDT850I2.Item("EDI_DTL_LINE_ID")
                            If IsDate(rowEDT850I1.Item("EDI_PO_DATE") & String.Empty) Then
                                rowSOTORDRX.Item("ORDR_DATE") = rowEDT850I1.Item("EDI_PO_DATE") & String.Empty
                            Else
                                rowSOTORDRX.Item("ORDR_DATE") = DateTime.Now.ToShortDateString
                            End If
                            rowSOTORDRX.Item("ORDR_CALLER_NAME") = "EDI"
                            rowSOTORDRX.Item("EDI_CUST_REF_NO") = TruncateField(rowEDT850I1.Item("EDI_CUST_REF_NO") & String.Empty, "SOTORDR1", "EDI_CUST_REF_NO")
                            rowSOTORDRX.Item("ORDR_LINE_SOURCE") = ORDR_SOURCE
                            rowSOTORDRX.Item("ORDR_TYPE_CODE") = "REG"
                            rowSOTORDRX.Item("ORDR_CUST_PO") = TruncateField(rowEDT850I1.Item("EDI_PO_ORDER_NO") & String.Empty, "SOTORDR1", "ORDR_CUST_PO")
                            rowSOTORDRX.Item("ORDR_COMMENT") = TruncateField((rowEDT850I1.Item("EDI_SPECIAL_INST") & String.Empty).ToString.Trim, "SOTORDR1", "ORDR_COMMENT")

                            ' Note: ST indicates Ship to Patient, 1T indicates the Ship To
                            If dst.Tables("EDT850I5").Select("EDI_ENTITY_ID_CODE = 'ST'").Length > 0 Then
                                rowSOTORDRX.Item("ORDR_DPD") = "1"
                            Else
                                rowSOTORDRX.Item("ORDR_DPD") = "0"
                            End If

                            ' Sold To
                            If dst.Tables("EDT850I5").Select("EDI_ENTITY_ID_CODE = 'BT'").Length <> 0 Then
                                CUST_CODE = dst.Tables("EDT850I5").Select("EDI_ENTITY_ID_CODE = 'BT'")(0).Item("EDI_ADDR_ID_CODE") & String.Empty
                                CUST_CODE = ABSolution.ASCMAIN1.Format_Field(CUST_CODE, "CUST_CODE")

                                If dst.Tables("EDT850I5").Select("EDI_ENTITY_ID_CODE = '1T'").Length > 0 Then
                                    CUST_SHIP_TO_NO = dst.Tables("EDT850I5").Select("EDI_ENTITY_ID_CODE = '1T'")(0).Item("EDI_ADDR_ID_CODE") & String.Empty
                                    CUST_SHIP_TO_NO = ABSolution.ASCMAIN1.Format_Field(CUST_SHIP_TO_NO, "CUST_SHIP_TO_NO") & String.Empty
                                End If

                                rowSOTORDRX.Item("CUST_CODE") = CUST_CODE
                                rowSOTORDRX.Item("CUST_SHIP_TO_NO") = CUST_SHIP_TO_NO

                                rowSOTORDRX.Item("CUST_NAME") = dst.Tables("EDT850I5").Select("EDI_ENTITY_ID_CODE = 'BT'")(0).Item("EDI_ADDR_NAME") & String.Empty

                            End If

                            rowARTCUST1 = baseClass.LookUp("ARTCUST1", CUST_CODE)
                            rowARTCUST2 = baseClass.LookUp("ARTCUST2", New String() {CUST_CODE, CUST_SHIP_TO_NO})
                            rowARTCUST3 = baseClass.LookUp("ARTCUST3", CUST_CODE)

                            If rowSOTORDRX.Item("ORDR_DPD") = "1" AndAlso ORDR_SOURCE = "E" Then
                                rowSOTORDRX.Item("ORDR_SHIP_COMPLETE") = "1"
                            ElseIf rowARTCUST1 IsNot Nothing Then
                                rowSOTORDRX.Item("ORDR_SHIP_COMPLETE") = rowARTCUST1.Item("CUST_SHIP_COMPLETE") & String.Empty
                            End If

                            rowSOTORDRX.Item("CUST_LINE_REF") = String.Empty
                            rowSOTORDRX.Item("ORDR_QTY") = Val(rowEDT850I2.Item("EDI_QTY_ORDERED") & String.Empty)

                            ITEM_UPC_CODE = (rowEDT850I2.Item("EDI_UPC") & String.Empty).ToString.Trim
                            If ITEM_UPC_CODE.Trim.Length > 0 Then
                                ITEM_UPC_CODE = ABSolution.ASCMAIN1.Format_Field(ITEM_UPC_CODE, "ITEM_UPC_CODE")
                            Else
                                ITEM_UPC_CODE = (rowEDT850I2.Item("EDI_PART_NUM_DESC") & String.Empty).ToString.Trim
                                rowEDT850I2.Item("EDI_UPC") = rowEDT850I2.Item("EDI_PART_NUM_DESC") & String.Empty
                                If ITEM_UPC_CODE.Length = 0 Then
                                    ITEM_UPC_CODE = "UNK"
                                    rowEDT850I2.Item("EDI_UPC") = "unk"
                                End If
                            End If

                            rowSOTORDRX.Item("ITEM_CODE") = ITEM_UPC_CODE
                            rowSOTORDRX.Item("ITEM_PROD_ID") = ITEM_UPC_CODE

                            Select Case rowEDT850I2.Item("EDI_LR_IND") & String.Empty
                                Case "OD"
                                    rowSOTORDRX.Item("ORDR_LR") = "R"
                                Case "OS"
                                    rowSOTORDRX.Item("ORDR_LR") = "L"
                            End Select

                            rowSOTORDRX.Item("ITEM_DESC") = TruncateField(rowEDT850I2.Item("EDI_ITEM_DESC_1") & String.Empty, "SOTORDR1", "ITEM_DESC")
                            rowSOTORDRX.Item("ITEM_DESC2") = TruncateField(rowEDT850I2.Item("EDI_ITEM_DESC_2") & String.Empty, "SOTORDR1", "ITEM_DESC2")

                            rowSOTORDRX.Item("SHIP_VIA_CODE") = String.Empty
                            If rowARTCUST2 IsNot Nothing Then
                                If rowSOTORDRX.Item("ORDR_DPD") = "1" AndAlso rowARTCUST2.Item("CUST_SHIP_TO_SHIP_VIA_CODE_DPD") & String.Empty <> String.Empty Then
                                    rowSOTORDRX.Item("SHIP_VIA_CODE") = rowARTCUST2.Item("CUST_SHIP_TO_SHIP_VIA_CODE_DPD") & String.Empty
                                Else
                                    rowSOTORDRX.Item("SHIP_VIA_CODE") = rowARTCUST2.Item("CUST_SHIP_TO_SHIP_VIA_CODE") & String.Empty
                                End If
                            End If

                            If rowSOTORDRX.Item("SHIP_VIA_CODE") & String.Empty = String.Empty AndAlso rowARTCUST1 IsNot Nothing Then
                                If rowARTCUST3 IsNot Nothing Then
                                    If rowSOTORDRX.Item("ORDR_DPD") = "1" AndAlso rowARTCUST3.Item("SHIP_VIA_CODE_DPD") & String.Empty <> String.Empty Then
                                        rowSOTORDRX.Item("SHIP_VIA_CODE") = rowARTCUST3.Item("SHIP_VIA_CODE_DPD") & String.Empty
                                    Else
                                        rowSOTORDRX.Item("SHIP_VIA_CODE") = rowARTCUST3.Item("SHIP_VIA_CODE") & String.Empty
                                    End If
                                End If
                            End If

                            If rowSOTORDRX.Item("ORDR_DPD") = "1" AndAlso rowSOTORDRX.Item("SHIP_VIA_CODE") & String.Empty = String.Empty Then
                                rowSOTORDRX.Item("SHIP_VIA_CODE") = DpdDefaultShipViaCode
                            End If

                            rowSOTORDRX.Item("ORDR_LOCK_SHIP_VIA") = "0"

                            PATIENT_NAME = String.Empty
                            If dst.Tables("EDT850I5").Select("EDI_ENTITY_ID_CODE = 'QC'").Length > 0 Then
                                PATIENT_NAME = dst.Tables("EDT850I5").Select("EDI_ENTITY_ID_CODE = 'QC'")(0).Item("EDI_ADDR_NAME") & String.Empty
                            End If
                            rowSOTORDRX.Item("PATIENT_NAME") = TruncateField(PATIENT_NAME, "SOTORDR2", "PATIENT_NAME")
                            rowSOTORDRX.Item("PATIENT_NAME") = StrConv(rowSOTORDRX.Item("PATIENT_NAME") & String.Empty, VbStrConv.ProperCase)

                            rowSOTORDRX.Item("ORDR_CUST_PO") = rowEDT850I1.Item("EDI_PO_ORDER_NO") & String.Empty

                            rowSOTORDRX.Item("CUST_PHONE") = ""

                            rowEDT850I5 = Nothing
                            If dst.Tables("EDT850I5").Select("EDI_ENTITY_ID_CODE = 'ST'").Length > 0 Then
                                rowEDT850I5 = dst.Tables("EDT850I5").Select("EDI_ENTITY_ID_CODE = 'ST'")(0)
                            ElseIf dst.Tables("EDT850I5").Select("EDI_ENTITY_ID_CODE = '1T'").Length > 0 Then
                                rowEDT850I5 = dst.Tables("EDT850I5").Select("EDI_ENTITY_ID_CODE = '1T'")(0)
                            End If

                            ' this is the address where the items will be shipped to
                            If rowEDT850I5 IsNot Nothing Then
                                strVal = (rowEDT850I5.Item("EDI_ADDR_NAME") & String.Empty).ToString.Replace("'", "").Trim

                                ' Fix the name
                                If rowSOTORDRX.Item("ORDR_DPD") & String.Empty = "1" Then
                                    strVal = StrConv(strVal, VbStrConv.ProperCase)
                                End If
                                rowSOTORDRX.Item("CUST_NAME") = TruncateField(strVal, "SOTORDR5", "CUST_NAME")

                                strVal = (rowEDT850I5.Item("EDI_ADDR_ADDRESS1") & String.Empty).ToString.Replace("'", "").Trim
                                rowSOTORDRX.Item("CUST_ADDR1") = TruncateField(strVal, "SOTORDR5", "CUST_ADDR1")

                                strVal = (rowEDT850I5.Item("EDI_ADDR_ADDRESS2") & String.Empty).ToString.Replace("'", "").Trim
                                rowSOTORDRX.Item("CUST_ADDR2") = TruncateField(strVal, "SOTORDR5", "CUST_ADDR2")

                                strVal = (rowEDT850I5.Item("EDI_ADDR_CITY") & String.Empty).ToString.Replace("'", "").Trim
                                rowSOTORDRX.Item("CUST_CITY") = TruncateField(strVal, "SOTORDR5", "CUST_CITY")

                                strVal = (rowEDT850I5.Item("EDI_ADDR_STATE") & String.Empty).ToString.Replace("'", "").Trim
                                rowSOTORDRX.Item("CUST_STATE") = TruncateField(strVal, "SOTORDR5", "CUST_STATE")

                                strVal = (rowEDT850I5.Item("EDI_ADDR_ZIP_CODE") & String.Empty).ToString.Trim
                                strVal = FormatZipCode(strVal)
                                rowSOTORDRX.Item("CUST_ZIP_CODE") = TruncateField(strVal, "SOTORDR5", "CUST_ZIP_CODE")
                                rowSOTORDRX.Item("CUST_COUNTRY") = "US"

                            End If

                            rowEDT850I5 = Nothing
                            If dst.Tables("EDT850I5").Select("EDI_ENTITY_ID_CODE = '1T'").Length > 0 Then
                                rowEDT850I5 = dst.Tables("EDT850I5").Select("EDI_ENTITY_ID_CODE = '1T'")(0)
                            End If

                            If rowEDT850I5 IsNot Nothing Then
                                strVal = (rowEDT850I5.Item("EDI_ADDR_NAME") & String.Empty).ToString.Replace("'", "").Trim
                                rowSOTORDRX.Item("CUST_SHIP_TO_NAME") = TruncateField(strVal, "SOTORDR5", "CUST_NAME")

                                strVal = (rowEDT850I5.Item("EDI_ADDR_ADDRESS1") & String.Empty).ToString.Replace("'", "").Trim
                                rowSOTORDRX.Item("CUST_SHIP_TO_ADDR1") = TruncateField(strVal, "SOTORDR5", "CUST_ADDR1")

                                strVal = (rowEDT850I5.Item("EDI_ADDR_ADDRESS2") & String.Empty).ToString.Replace("'", "").Trim
                                rowSOTORDRX.Item("CUST_SHIP_TO_ADDR2") = TruncateField(strVal, "SOTORDR5", "CUST_ADDR2")

                                strVal = (rowEDT850I5.Item("EDI_ADDR_CITY") & String.Empty).ToString.Replace("'", "").Trim
                                rowSOTORDRX.Item("CUST_SHIP_TO_CITY") = TruncateField(strVal, "SOTORDR5", "CUST_CITY")

                                strVal = (rowEDT850I5.Item("EDI_ADDR_STATE") & String.Empty).ToString.Replace("'", "").Trim
                                rowSOTORDRX.Item("CUST_SHIP_TO_STATE") = TruncateField(strVal, "SOTORDR5", "CUST_STATE")

                                strVal = (rowEDT850I5.Item("EDI_ADDR_ZIP_CODE") & String.Empty).ToString.Trim
                                strVal = FormatZipCode(strVal)
                                rowSOTORDRX.Item("CUST_SHIP_TO_ZIP_CODE") = TruncateField(strVal, "SOTORDR5", "CUST_ZIP_CODE")
                                rowSOTORDRX.Item("CUST_SHIP_TO_COUNTRY") = "US"
                            End If

                            If rowARTCUST1 IsNot Nothing Then
                                rowSOTORDRX.Item("BILLING_NAME") = TruncateField(rowARTCUST1.Item("CUST_NAME") & String.Empty, "SOTORDR5", "CUST_NAME")
                                rowSOTORDRX.Item("BILLING_ADDRESS1") = TruncateField(rowARTCUST1.Item("CUST_ADDR1") & String.Empty, "SOTORDR5", "CUST_ADDR1")
                                rowSOTORDRX.Item("BILLING_ADDRESS2") = TruncateField(rowARTCUST1.Item("CUST_ADDR2") & String.Empty, "SOTORDR5", "CUST_ADDR2")
                                rowSOTORDRX.Item("BILLING_CITY") = TruncateField(rowARTCUST1.Item("CUST_CITY") & String.Empty, "SOTORDR5", "CUST_CITY")
                                rowSOTORDRX.Item("BILLING_STATE") = TruncateField(rowARTCUST1.Item("CUST_STATE") & String.Empty, "SOTORDR5", "CUST_STATE")
                                rowSOTORDRX.Item("BILLING_ZIP") = TruncateField(rowARTCUST1.Item("CUST_ZIP_CODE") & String.Empty, "SOTORDR5", "CUST_ZIP_CODE")
                            End If

                            If rowSOTORDRX.Item("ORDR_DPD") & String.Empty <> "1" Then
                                If CUST_SHIP_TO_NO.Length > 0 AndAlso rowARTCUST2 IsNot Nothing Then
                                    rowSOTORDRX.Item("CUST_NAME") = rowARTCUST2.Item("CUST_SHIP_TO_NAME") & String.Empty
                                ElseIf rowARTCUST1 IsNot Nothing Then
                                    rowSOTORDRX.Item("CUST_NAME") = rowARTCUST1.Item("CUST_NAME") & String.Empty
                                End If
                            End If

                            dst.Tables("SOTORDRX").Rows.Add(rowSOTORDRX)
                        Next
                    Next
                Next

                If dst.Tables("SOTORDRX").Rows.Count > 0 Then
                    ' Commit the data from the Excel file and then archive the file
                    Dim UpdateInProcess As Boolean = False
                    With baseClass
                        Try
                            .BeginTrans()
                            UpdateInProcess = True
                            .clsASCBASE1.Update_Record_TDA("SOTORDRX")
                            .clsASCBASE1.Update_Record_TDA("EDT850I0")
                            .clsASCBASE1.Update_Record_TDA("EDT850I1")
                            .clsASCBASE1.Update_Record_TDA("EDT850I2")
                            .clsASCBASE1.Update_Record_TDA("EDT850I5")
                            .CommitTrans()
                            UpdateInProcess = False
                        Catch ex As Exception
                            If UpdateInProcess Then .Rollback()
                            RecordLogEntry("ProcessEDISalesOrders: " & ex.Message)
                        End Try

                    End With

                End If

                Dim ORDR_NO As String = String.Empty
                Dim ORDR_CALLER_NAME As String = String.Empty
                Dim ORDR_SHIP_COMPLETE As String = String.Empty
                Dim createShipTo As Boolean = ORDR_SOURCE = "E"
                Dim locateShipTobyTelephoneNo As Boolean = False
                Dim alwaysUseImportedShipToAddress As Boolean = True

                dst.Tables("SOTORDRX").Rows.Clear()

                ' Need to process each order individually for pricing reasons; therefore
                ' need to move the data to a temp data table and process each order individually
                For Each headers As DataRow In ABSolution.ASCDATA1.GetDataTable("SELECT DISTINCT ORDR_NO FROM SOTORDRX WHERE PROCESS_IND IS NULL AND ORDR_SOURCE = :PARM1", String.Empty, "V", New Object() {ORDR_SOURCE}).Rows
                    ClearDataSetTables(True)
                    ORDR_NO = headers.Item("ORDR_NO") & String.Empty
                    baseClass.clsASCBASE1.Fill_Records("SOTORDRX", New Object() {ORDR_SOURCE, ORDR_NO})

                    If dst.Tables("SOTORDRX").Rows.Count = 0 Then
                        RecordLogEntry("ProcessEDISalesOrders: Invalid Order Number (" & ORDR_NO & ") for Source " & ORDR_SOURCE)
                        Continue For
                    End If

                    ORDR_NO = String.Empty
                    ORDR_CALLER_NAME = dst.Tables("SOTORDRX").Rows(0).Item("ORDR_CALLER_NAME") & String.Empty
                    ORDR_SHIP_COMPLETE = dst.Tables("SOTORDRX").Rows(0).Item("ORDR_SHIP_COMPLETE") & String.Empty
                    If CreateSalesOrder(ORDR_NO, createShipTo, locateShipTobyTelephoneNo, ORDR_SOURCE, ORDR_SOURCE, ORDR_CALLER_NAME, alwaysUseImportedShipToAddress, ORDR_SOURCE) Then
                        dst.Tables("SOTORDR1").Rows(0).Item("ORDR_SHIP_COMPLETE") = ORDR_SHIP_COMPLETE
                        UpdateDataSetTables()
                        salesOrdersProcessed += 1
                        Try
                            ABSolution.ASCDATA1.ExecuteSQL("DELETE FROM SOTORDRX WHERE ORDR_SOURCE = :PARM1 AND ORDR_NO = :PARM2", _
                                                           "VV", _
                                                        New Object() {ORDR_SOURCE, dst.Tables("SOTORDRX").Rows(0).Item("ORDR_NO") & String.Empty})

                        Catch ex As Exception
                            RecordLogEntry("ProcessEDISalesOrders: " & "Could not create sales order for " & dst.Tables("SOTORDRX").Rows(0).Item("ORDR_NO") & ":" & ex.Message)
                        End Try
                    Else
                        RecordLogEntry("ProcessEDISalesOrders: " & "Could not create sales order for " & dst.Tables("SOTORDRX").Rows(0).Item("ORDR_NO"))
                    End If
                Next
            Catch ex As Exception
                RecordLogEntry("ProcessEDISalesOrders: " & ex.Message)
                emailErrors(ORDR_SOURCE, 1, ex.Message)
            Finally
                RecordLogEntry(salesOrdersProcessed & " " & ediConnection.ConnectionDescription & " Sales Orders imported.")
            End Try


        End Sub

        Private Sub ProcessBLScan(ByVal ORDR_SOURCE As String)

            Dim ftpConnection As New Connection(ORDR_SOURCE)
            Dim orderFileName As String = String.Empty
            Dim orderData As String = String.Empty
            Dim orderElements() As String
            Dim tempStr As String = String.Empty
            Dim rowSOTORDRX As DataRow = Nothing

            Dim CUST_CODE As String = String.Empty
            Dim CUST_SHIP_TO_NO As String = String.Empty
            Dim ORDR_LR As String = String.Empty
            Dim custShip As String = String.Empty
            Dim salesOrdersProcessed As Integer = 0
            Dim ORDR_LNO As Integer = 0
            Dim ORDR_NO As String = String.Empty
            Dim recCount As Integer = 1

            ' Tweak the order source
            ORDR_SOURCE = "E"

            ' Perform FTP Here
            ftpFileList.Clear()
            ImportedFiles.Clear()

            Try
                Ftp1.RuntimeLicense = nSoftwareftpkey
                Ftp1.User = ftpConnection.UserId
                Ftp1.Password = ftpConnection.Password
                Ftp1.RemoteHost = ftpConnection.RemoteHost
                Ftp1.Logon()

                ftpFileList = New List(Of String)
                If Ftp1.Connected Then
                    If ftpConnection.RemoteInDirectory.Length > 0 Then
                        Ftp1.RemotePath = "/" & ftpConnection.RemoteInDirectory
                    End If
                    Ftp1.ListDirectory()
                End If

                For Each fileFtp As String In ftpFileList

                    If fileFtp.Length = 0 Then Continue For
                    If Not fileFtp.ToUpper.Trim.EndsWith(".CSV") Then Continue For

                    Ftp1.Overwrite = True
                    Ftp1.RemoteFile = fileFtp
                    Ftp1.LocalFile = ftpConnection.LocalInDir & fileFtp
                    Ftp1.Download()
                    Ftp1.DeleteFile(fileFtp)

                Next

            Catch ex As Exception
                RecordLogEntry("ProcessBLScan: (ftp) " & ex.Message)
            Finally
                Ftp1.Logoff()
                Ftp1.Dispose()
            End Try

            If testMode Then
                RecordLogEntry("ProcessBLScan: ftpConnection.LocalInDir = " & ftpConnection.LocalInDir)
            End If

            Try
                For Each orderFile As String In My.Computer.FileSystem.GetFiles(ftpConnection.LocalInDir, FileIO.SearchOption.SearchTopLevelOnly, "*.csv")

                    orderFileName = My.Computer.FileSystem.GetName(orderFile)
                    RecordLogEntry("Importing file: " & orderFileName)

                    Using orderReader As New StreamReader(orderFile)
                        System.Threading.Thread.Sleep(1000)
                        recCount += 1
                        ORDR_NO = "BL" & recCount.ToString.Trim & DateTime.Now.ToString("hhmmss")
                        ORDR_LNO = 0

                        ' The first record will be column descriptions and can be ignored
                        If orderReader.Peek <> -1 Then
                            orderData = orderReader.ReadLine()
                        End If

                        While orderReader.Peek <> -1

                            orderData = orderReader.ReadLine()
                            orderElements = orderData.Split(",")

                            If orderElements.Length <> 33 AndAlso orderElements.Length <> 34 Then
                                RecordLogEntry("ProcessBLScan: Order file invalid (" & orderFileName & ")")
                                Continue While
                            End If

                            rowSOTORDRX = dst.Tables("SOTORDRX").NewRow
                            rowSOTORDRX.Item("ORDR_NO") = ORDR_NO
                            ORDR_LNO += 1
                            rowSOTORDRX.Item("ORDR_LNO") = ORDR_LNO
                            rowSOTORDRX.Item("ORDR_SOURCE") = ORDR_SOURCE

                            custShip = (orderElements(7) & String.Empty).ToString.Trim
                            CUST_CODE = String.Empty
                            CUST_SHIP_TO_NO = String.Empty

                            Select Case custShip.Length
                                Case Is <= 6
                                    CUST_CODE = custShip
                                Case Else
                                    CUST_CODE = custShip.Substring(0, 6).Trim
                                    CUST_SHIP_TO_NO = custShip.Substring(6).Trim
                            End Select

                            If CUST_SHIP_TO_NO.Length > 6 Then
                                CUST_SHIP_TO_NO = CUST_SHIP_TO_NO.Substring(0, 6).Trim
                            End If

                            CUST_CODE = ABSolution.ASCMAIN1.Format_Field(CUST_CODE, "CUST_CODE")
                            If CUST_SHIP_TO_NO.Length > 0 Then
                                CUST_SHIP_TO_NO = ABSolution.ASCMAIN1.Format_Field(CUST_SHIP_TO_NO, "CUST_SHIP_TO_NO")
                            End If

                            If CUST_SHIP_TO_NO = "000000" Then
                                CUST_SHIP_TO_NO = String.Empty
                            End If

                            rowSOTORDRX.Item("CUST_CODE") = CUST_CODE
                            rowSOTORDRX.Item("CUST_SHIP_TO_NO") = CUST_SHIP_TO_NO
                            rowSOTORDRX.Item("CUST_NAME") = orderElements(22) & String.Empty

                            rowSOTORDRX.Item("ORDR_TYPE_CODE") = "REG"
                            rowSOTORDRX.Item("ORDR_DPD") = "0"
                            rowSOTORDRX.Item("ORDR_QTY") = Val(orderElements(17) & String.Empty)
                            rowSOTORDRX.Item("ITEM_CODE") = TruncateField(orderElements(11), "SOTORDR2", "ITEM_CODE")
                            If (rowSOTORDRX.Item("ITEM_CODE") & String.Empty).ToString.Length > 12 Then
                                rowSOTORDRX.Item("ITEM_CODE") = rowSOTORDRX.Item("ITEM_CODE").ToString.Substring(0, 12).Trim
                            End If

                            ' Validate UOM
                            If (rowSOTORDRX.Item("ITEM_CODE") & String.Empty).ToString.Length > 0 Then
                                rowSOTORDRX.Item("ITEM_CODE") = ABSolution.ASCMAIN1.Format_Field(rowSOTORDRX.Item("ITEM_CODE"), "ITEM_UPC_CODE")
                                rowICTITEM1 = ABSolution.ASCDATA1.GetDataRow("SELECT * FROM ICTITEM1 WHERE ITEM_UPC_CODE = :PARM1", "V", New Object() {rowSOTORDRX.Item("ITEM_CODE")})
                            Else
                                rowICTITEM1 = Nothing
                            End If

                            If rowICTITEM1 IsNot Nothing Then
                                If (rowICTITEM1.Item("ITEM_UOM") & String.Empty).ToString.Trim.ToUpper <> _
                                    orderElements(18).Trim.ToUpper Then
                                    rowSOTORDRX.Item("ITEM_UOM") = "1"
                                End If
                            End If

                            rowARTCUST1 = baseClass.LookUp("ARTCUST1", CUST_CODE)
                            rowARTCUST2 = baseClass.LookUp("ARTCUST2", New String() {CUST_CODE, CUST_SHIP_TO_NO})
                            rowARTCUST3 = baseClass.LookUp("ARTCUST3", CUST_CODE)

                            ORDR_LR = String.Empty
                            If ORDR_LR.Length > 2 Then ORDR_LR = ORDR_LR.Substring(0, 2)
                            Select Case ORDR_LR
                                Case "OD"
                                    rowSOTORDRX.Item("ORDR_LR") = "R"
                                Case "OS"
                                    rowSOTORDRX.Item("ORDR_LR") = "L"
                                Case "OU"
                                    rowSOTORDRX.Item("ORDR_LR") = "B"
                            End Select
                            rowSOTORDRX.Item("ITEM_DESC") = TruncateField(orderElements(28), "SOTORDR2", "ITEM_DESC")

                            If IsDate(orderElements(2).Trim) Then
                                rowSOTORDRX.Item("ORDR_DATE") = CDate(orderElements(2).Trim).ToString("dd-MMM-yyyy")
                            Else
                                rowSOTORDRX.Item("ORDR_DATE") = DateTime.Now.ToString("dd-MMM-yyyy")
                            End If

                            If rowARTCUST3 IsNot Nothing Then
                                rowSOTORDRX.Item("SHIP_VIA_CODE") = rowARTCUST3.Item("SHIP_VIA_CODE")
                            End If

                            ' Overnight delivery
                            If orderElements(5).Trim.ToUpper = "ON" Then
                                rowSOTSVIA1 = baseClass.LookUp("SOTSVIA1", rowSOTORDRX.Item("SHIP_VIA_CODE") & String.Empty)
                                If rowSOTSVIA1 IsNot Nothing AndAlso (rowSOTSVIA1.Item("SHIP_VIA_COD_IND") & String.Empty).ToString.Trim = "1" Then
                                    rowSOTORDRX.Item("SHIP_VIA_CODE") = SO_PARM_SHIP_ND_COD
                                ElseIf Me.SO_PARM_SHIP_ND.Length > 0 Then
                                    rowSOTORDRX.Item("SHIP_VIA_CODE") = SO_PARM_SHIP_ND
                                End If
                            End If
                            rowSOTORDRX.Item("ORDR_LOCK_SHIP_VIA") = "0"

                            If rowARTCUST1 IsNot Nothing Then
                                rowSOTORDRX.Item("ORDR_SHIP_COMPLETE") = Val(rowARTCUST1.Item("CUST_SHIP_COMPLETE") & String.Empty)
                            End If

                            rowSOTORDRX.Item("PATIENT_NAME") = orderElements(26)
                            rowSOTORDRX.Item("PATIENT_NAME") = TruncateField(rowSOTORDRX.Item("PATIENT_NAME") & String.Empty, "SOTORDR2", "PATIENT_NAME")
                            rowSOTORDRX.Item("PATIENT_NAME") = StrConv(rowSOTORDRX.Item("PATIENT_NAME") & String.Empty, VbStrConv.ProperCase)

                            orderElements(1) = orderElements(1).Replace("'", "")
                            rowSOTORDRX.Item("ORDR_CUST_PO") = TruncateField(orderElements(1), "SOTORDR1", "ORDR_CUST_PO")
                            If (rowSOTORDRX.Item("ORDR_CUST_PO") & String.Empty).ToString.Length = 0 Then
                                rowSOTORDRX.Item("ORDR_CUST_PO") = TruncateField("BLSCAN ORDER", "SOTORDR1", "ORDR_CUST_PO")
                            End If

                            dst.Tables("SOTORDRX").Rows.Add(rowSOTORDRX)
                        End While
                    End Using

                    ImportedFiles.Add(orderFile)
                Next

                If dst.Tables("SOTORDRX").Rows.Count > 0 Then
                    ' Commit the data from the Excel file and then archive the file
                    Dim UpdateInProcess As Boolean = False
                    With baseClass
                        Try
                            .BeginTrans()
                            UpdateInProcess = True
                            .clsASCBASE1.Update_Record_TDA("SOTORDRX")
                            .CommitTrans()
                            UpdateInProcess = False

                        Catch ex As Exception
                            If UpdateInProcess Then .Rollback()
                            RecordLogEntry("ProcessBLScan: " & ex.Message)
                        End Try

                    End With

                End If

                ' If DPD then set the LR
                ORDR_LR = String.Empty
                ORDR_NO = String.Empty
                Dim ORDR_CALLER_NAME As String = String.Empty
                Dim ORDR_SHIP_COMPLETE As String = String.Empty
                dst.Tables("SOTORDRX").Rows.Clear()

                ' Need to process each order individually for pricing reasons; therefore
                ' need to move the datat to a temp data table and process each order individually
                For Each headers As DataRow In ABSolution.ASCDATA1.GetDataTable("SELECT DISTINCT ORDR_NO FROM SOTORDRX WHERE PROCESS_IND IS NULL AND ORDR_SOURCE = :PARM1", String.Empty, "V", New Object() {ORDR_SOURCE}).Rows
                    ClearDataSetTables(True)
                    ORDR_NO = headers.Item("ORDR_NO") & String.Empty
                    baseClass.clsASCBASE1.Fill_Records("SOTORDRX", New Object() {ORDR_SOURCE, ORDR_NO})

                    If dst.Tables("SOTORDRX").Rows.Count = 0 Then
                        RecordLogEntry("ProcessBLScan: Invalid Order Number (" & ORDR_NO & ") for " & ftpConnection.ConnectionDescription)
                        Continue For
                    End If

                    'ORDR_NO = String.Empty
                    ORDR_CALLER_NAME = dst.Tables("SOTORDRX").Rows(0).Item("ORDR_CALLER_NAME") & String.Empty
                    ORDR_SHIP_COMPLETE = dst.Tables("SOTORDRX").Rows(0).Item("ORDR_SHIP_COMPLETE") & String.Empty
                    If CreateSalesOrder(ORDR_NO, False, False, ORDR_SOURCE, ORDR_SOURCE, ORDR_CALLER_NAME, False, "B") Then
                        dst.Tables("SOTORDR1").Rows(0).Item("ORDR_SHIP_COMPLETE") = ORDR_SHIP_COMPLETE
                        UpdateDataSetTables()
                        salesOrdersProcessed += 1
                        Try
                            ABSolution.ASCDATA1.ExecuteSQL("DELETE FROM SOTORDRX WHERE ORDR_SOURCE = :PARM1 AND ORDR_NO = :PARM2", _
                                                           "VV", _
                                                        New Object() {ORDR_SOURCE, dst.Tables("SOTORDRX").Rows(0).Item("ORDR_NO") & String.Empty})

                        Catch ex As Exception

                        End Try
                    Else
                        RecordLogEntry("ProcessBLScan: " & "Could not create sales order for " & dst.Tables("SOTORDRX").Rows(0).Item("ORDR_NO"))
                    End If
                Next
            Catch ex As Exception
                RecordLogEntry("ProcessBLScan: " & ex.Message)
                emailErrors(ORDR_SOURCE, 1, ex.Message)
            Finally
                RecordLogEntry(salesOrdersProcessed & " " & ftpConnection.ConnectionDescription & " Sales Orders imported.")

                ' Move CSN and any SNT file extensions to the archive directory
                For Each orderFile As String In ImportedFiles
                    My.Computer.FileSystem.MoveFile(orderFile, ftpConnection.LocalInDirArchive & My.Computer.FileSystem.GetName(orderFile), True)
                Next

            End Try

        End Sub

        Private Sub ProcessVisionWebSalesOrders(ByVal ORDR_SOURCE As String)

            Dim vwConnection As New Connection(ORDR_SOURCE)
            Dim rowSOTORDRX As DataRow = Nothing
            Dim ORDR_NO As String = String.Empty
            Dim ORDR_LNO As Integer = 0
            Dim rowData As DataRow = Nothing
            Dim Telephone As String = String.Empty
            Dim AttentionTo As String = String.Empty
            Dim orderNumLoop As Int16 = 0
            Dim salesordersprocessed As Int16 = 0

            Dim SOFT_Id As String = String.Empty
            Dim HEADER_Id As String = String.Empty
            Dim CUSTOMER_Id As String = String.Empty
            Dim ACCOUNTS_Id As String = String.Empty
            Dim ACCOUNT_Id As String = String.Empty

            Dim DELIVERY_Id As String = String.Empty

            Dim SOFTCONTACTS_id As String = String.Empty
            Dim SOFTCONTACT_id As String = String.Empty
            Dim PRESCRIPTION_id As String = String.Empty

            Dim CUST_CODE As String = String.Empty
            Dim CUST_SHIP_TO_NO As String = String.Empty
            Dim creationDate As String = String.Empty

            Dim itemCode As String = String.Empty
            Dim upcCode As String = String.Empty
            Dim fileprocessing As String = String.Empty

            Try
                ImportedFiles.Clear()

                Dim xsdFile As String = vwConnection.LocalInDir & "ContactLensOrder.XSD"

                If Not My.Computer.FileSystem.FileExists(xsdFile) Then
                    RecordLogEntry("ProcessVisionWebSalesOrders: " & xsdFile & " could not be found.")
                    Exit Sub
                End If

                ' Create the DataSet to read the schema into.
                vwXmlDataset = New DataSet
                'Create a FileStream object with the file path and name.
                Dim myFileStream As System.IO.FileStream = New System.IO.FileStream(xsdFile, System.IO.FileMode.Open)
                'Create a new XmlTextReader object with the FileStream.
                Dim myXmlTextReader As System.Xml.XmlTextReader = New System.Xml.XmlTextReader(myFileStream)
                'Read the schema into the DataSet and close the reader.
                vwXmlDataset.ReadXmlSchema(myXmlTextReader)
                myXmlTextReader.Close()
                Dim errorfile As Boolean = False

                For Each orderFile As String In My.Computer.FileSystem.GetFiles(vwConnection.LocalInDir, FileIO.SearchOption.SearchTopLevelOnly, "*.xml")

                    fileprocessing = orderFile
                    errorfile = False
                    For Each tbl As DataTable In vwXmlDataset.Tables
                        tbl.Rows.Clear()
                        tbl.BeginLoadData()
                    Next

                    vwXmlDataset.ReadXml(orderFile)

                    For Each tbl As DataTable In vwXmlDataset.Tables
                        Try
                            tbl.EndLoadData()
                        Catch ex As Exception
                            ImportedFiles.Add(orderFile)
                            RecordLogEntry("Error loading VW file " & orderFile & ": " & ex.Message)
                            errorfile = True
                            emailErrors(ORDR_SOURCE, 1, "Error loading VW file " & orderFile & ": " & ex.Message)
                            Exit For
                        End Try
                    Next

                    If errorfile Then Continue For

                    SOFT_Id = String.Empty

                    ' These are the 2 types of Orders in the XML Document
                    For Each tableName As String In New String() {"STK_SOFT_OFFICE", "RX_SOFT_PATIENT"}
                        For Each rowSoftType As DataRow In vwXmlDataset.Tables(tableName).Select("", tableName & "_Id")
                            SOFT_Id = rowSoftType.Item(tableName & "_Id") & String.Empty
                            orderNumLoop += 1
                            ORDR_NO = rowSoftType.Item("Id") & orderNumLoop.ToString
                            If ORDR_NO.Length > 10 Then ORDR_NO = ORDR_NO.Substring(0, 10).Trim

                            rowSOTORDRX = dst.Tables("SOTORDRX").NewRow
                            rowSOTORDRX.Item("ORDR_NO") = ORDR_NO
                            rowSOTORDRX.Item("ORDR_SOURCE") = ORDR_SOURCE

                            creationDate = (rowSoftType.Item("Creation_date") & String.Empty).ToString.Replace("T", Space(1))
                            If IsDate(creationDate) Then
                                rowSOTORDRX.Item("ORDR_DATE") = CDate(creationDate).ToString("dd-MMM-yyyy")
                            Else
                                rowSOTORDRX.Item("ORDR_DATE") = DateTime.Now
                            End If

                            rowSOTORDRX.Item("ORDR_COMMENT") = TruncateField(rowSoftType.Item("Comment") & String.Empty, "SOTORDRX", "ORDR_COMMENT")
                            rowSOTORDRX.Item("EDI_CUST_REF_NO") = rowSoftType.Item("Id") & String.Empty
                            rowSOTORDRX.Item("ORDR_TYPE_CODE") = "REG"
                            rowSOTORDRX.Item("ORDR_LNO") = 1
                            dst.Tables("SOTORDRX").Rows.Add(rowSOTORDRX)

                            HEADER_Id = String.Empty
                            CUSTOMER_Id = String.Empty
                            ACCOUNTS_Id = String.Empty
                            ACCOUNT_Id = String.Empty

                            DELIVERY_Id = String.Empty
                            SOFTCONTACTS_id = String.Empty
                            SOFTCONTACT_id = String.Empty
                            PRESCRIPTION_id = String.Empty

                            For Each rowHeader As DataRow In vwXmlDataset.Tables("HEADER").Select(tableName & "_Id = " & SOFT_Id, "HEADER_Id")
                                HEADER_Id = rowHeader.Item("HEADER_Id") & String.Empty
                                rowSOTORDRX.Item("ORDR_CUST_PO") = TruncateField(rowSoftType.Item("Id") & IIf(rowHeader.Item("PurchaseOrderNumber") & String.Empty <> String.Empty, ":" & rowHeader.Item("PurchaseOrderNumber"), String.Empty), "SOTORDRX", "ORDR_CUST_PO")
                                rowSOTORDRX.Item("ORDR_CALLER_NAME") = TruncateField(rowHeader.Item("OrderPlacedBy") & String.Empty, "SOTORDRX", "ORDR_CALLER_NAME")

                                CUSTOMER_Id = String.Empty
                                For Each rowCustomer As DataRow In vwXmlDataset.Tables("CUSTOMER").Select("HEADER_Id = " & HEADER_Id, "CUSTOMER_ID")
                                    CUSTOMER_Id = rowCustomer.Item("CUSTOMER_Id") & String.Empty

                                    For Each rowACCOUNTS As DataRow In vwXmlDataset.Tables("ACCOUNTS").Select("CUSTOMER_Id = " & CUSTOMER_Id, "ACCOUNTS_ID")
                                        ACCOUNTS_Id = rowACCOUNTS.Item("ACCOUNTS_ID") & String.Empty

                                        For Each rowACCOUNT As DataRow In vwXmlDataset.Tables("ACCOUNT").Select("ACCOUNTS_ID = " & ACCOUNTS_Id, "ACCOUNT_ID")

                                            Select Case (rowACCOUNT.Item("Class") & String.Empty).ToString.Trim.ToUpper
                                                Case "BIL"
                                                    ACCOUNT_Id = rowACCOUNT.Item("ACCOUNT_Id") & String.Empty
                                                    If ACCOUNT_Id.Length > 0 AndAlso vwXmlDataset.Tables("ADDRESS").Select("ACCOUNT_Id = " & ACCOUNT_Id).Length > 0 Then
                                                        rowData = vwXmlDataset.Tables("ADDRESS").Select("ACCOUNT_Id = " & ACCOUNT_Id)(0)
                                                        rowSOTORDRX.Item("CUST_SHIP_TO_NAME") = String.Empty
                                                        rowSOTORDRX.Item("CUST_SHIP_TO_ADDR1") = TruncateField(((rowData.Item("Street_Number") & String.Empty).ToString.Trim & " " & (rowData.Item("Street_Name") & String.Empty).ToString.Trim), "SOTORDRX", "CUST_SHIP_TO_ADDR1")
                                                        rowSOTORDRX.Item("CUST_SHIP_TO_ADDR2") = TruncateField((rowData.Item("Suite") & String.Empty).ToString.Trim, "SOTORDRX", "CUST_SHIP_TO_ADDR2")
                                                        rowSOTORDRX.Item("CUST_SHIP_TO_ADDR3") = String.Empty
                                                        rowSOTORDRX.Item("CUST_SHIP_TO_CITY") = TruncateField((rowData.Item("City") & String.Empty).ToString.Trim, "SOTORDRX", "CUST_SHIP_TO_CITY")
                                                        rowSOTORDRX.Item("CUST_SHIP_TO_STATE") = TruncateField((rowData.Item("State") & String.Empty).ToString.Trim, "SOTORDRX", "CUST_SHIP_TO_STATE")
                                                        rowSOTORDRX.Item("CUST_SHIP_TO_ZIP_CODE") = TruncateField((rowData.Item("ZipCode") & String.Empty).ToString.Trim, "SOTORDRX", "CUST_SHIP_TO_ZIP_CODE")
                                                        Telephone = (rowData.Item("TEL") & String.Empty).ToString.Trim
                                                        Telephone = FormatTelePhone(Telephone)
                                                        rowSOTORDRX.Item("CUST_SHIP_TO_PHONE") = TruncateField(Telephone, "SOTORDRX", "CUST_SHIP_TO_PHONE")
                                                        rowSOTORDRX.Item("CUST_SHIP_TO_COUNTRY") = TruncateField((rowData.Item("Country") & String.Empty).ToString.Trim, "SOTORDRX", "CUST_SHIP_TO_COUNTRY")
                                                        rowSOTORDRX.Item("CUST_SHIP_TO_FAX") = String.Empty
                                                        rowSOTORDRX.Item("CUST_SHIP_TO_EMAIL") = String.Empty
                                                    End If

                                                Case "SHP"
                                                    CUST_CODE = String.Empty
                                                    CUST_SHIP_TO_NO = String.Empty
                                                    ACCOUNT_Id = rowACCOUNT.Item("ACCOUNT_Id") & String.Empty
                                                    If ACCOUNT_Id.Length > 0 AndAlso vwXmlDataset.Tables("ADDRESS").Select("ACCOUNT_Id = " & ACCOUNT_Id).Length > 0 Then
                                                        rowData = vwXmlDataset.Tables("ADDRESS").Select("ACCOUNT_Id = " & ACCOUNT_Id)(0)
                                                        rowSOTORDRX.Item("CUST_NAME") = String.Empty
                                                        rowSOTORDRX.Item("CUST_ADDR1") = TruncateField(((rowData.Item("Street_Number") & String.Empty).ToString.Trim & " " & rowData.Item("Street_Name") & String.Empty).ToString.Trim, "SOTORDRX", "CUST_ADDR1")
                                                        rowSOTORDRX.Item("CUST_ADDR2") = TruncateField((rowData.Item("Suite") & String.Empty).ToString.Trim, "SOTORDRX", "CUST_ADDR2")
                                                        rowSOTORDRX.Item("CUST_CITY") = TruncateField((rowData.Item("City") & String.Empty).ToString.Trim, "SOTORDRX", "CUST_CITY")
                                                        rowSOTORDRX.Item("CUST_STATE") = TruncateField((rowData.Item("State") & String.Empty).ToString.Trim, "SOTORDRX", "CUST_STATE")
                                                        rowSOTORDRX.Item("CUST_ZIP_CODE") = TruncateField((rowData.Item("Zipcode") & String.Empty).ToString.Trim, "SOTORDRX", "CUST_ZIP_CODE")
                                                        Telephone = (rowData.Item("TEL") & String.Empty).ToString.Trim
                                                        Telephone = FormatTelePhone(Telephone)
                                                        rowSOTORDRX.Item("CUST_PHONE") = TruncateField(Telephone, "SOTORDRX", "CUST_PHONE")
                                                        rowSOTORDRX.Item("CUST_FAX") = String.Empty
                                                        rowSOTORDRX.Item("CUST_EMAIL") = String.Empty
                                                        rowSOTORDRX.Item("CUST_COUNTRY") = TruncateField((rowData.Item("Country") & String.Empty).ToString.Trim, "SOTORDRX", "CUST_COUNTRY")

                                                        CUST_CODE = rowACCOUNT.Item("Name") & String.Empty
                                                        If CUST_CODE.Contains("-") Then
                                                            CUST_SHIP_TO_NO = Split(CUST_CODE, "-")(1)
                                                            CUST_CODE = Split(CUST_CODE, "-")(0)
                                                        End If

                                                        CUST_CODE = ABSolution.ASCMAIN1.Format_Field(CUST_CODE, "CUST_CODE")
                                                        If CUST_CODE.Length = 0 Then
                                                            CUST_CODE = rowACCOUNT.Item("Name") & String.Empty
                                                        End If
                                                        CUST_CODE = TruncateField(CUST_CODE, "SOTORDRX", "CUST_CODE")
                                                        If CUST_SHIP_TO_NO.Length > 0 Then
                                                            CUST_SHIP_TO_NO = ABSolution.ASCMAIN1.Format_Field(CUST_SHIP_TO_NO, "CUST_SHIP_TO_NO")
                                                            CUST_SHIP_TO_NO = TruncateField(CUST_SHIP_TO_NO, "SOTORDRX", "CUST_SHIP_TO_NO")
                                                        End If

                                                        rowSOTORDRX.Item("CUST_CODE") = CUST_CODE
                                                        rowSOTORDRX.Item("CUST_SHIP_TO_NO") = CUST_SHIP_TO_NO
                                                    End If
                                            End Select
                                        Next
                                    Next
                                Next

                                ' See if DPD Order
                                rowSOTORDRX.Item("ORDR_DPD") = IIf(tableName = "RX_SOFT_PATIENT", "1", "0")
                                AttentionTo = String.Empty
                                DELIVERY_Id = String.Empty
                                If vwXmlDataset.Tables("DELIVERY").Select("HEADER_Id = " & HEADER_Id, "DELIVERY_Id").Length > 0 Then
                                    rowData = vwXmlDataset.Tables("DELIVERY").Select("HEADER_Id = " & HEADER_Id, "DELIVERY_Id")(0)
                                    DELIVERY_Id = rowData.Item("DELIVERY_Id") & String.Empty
                                    AttentionTo = (rowData.Item("AttentionTo") & String.Empty).ToString.Trim

                                    If vwXmlDataset.Tables("DELIVERY_METHOD").Select("DELIVERY_id = " & DELIVERY_Id).Length > 0 Then
                                        rowData = vwXmlDataset.Tables("DELIVERY_METHOD").Select("DELIVERY_id = " & DELIVERY_Id)(0)
                                        rowSOTORDRX.Item("SHIP_VIA_CODE") = (rowData.Item("Name") & String.Empty).ToString.Trim

                                        If (rowSOTORDRX.Item("SHIP_VIA_CODE") & String.Empty).ToString.ToUpper = "STANDARD CONTRACT" Then
                                            rowSOTORDRX.Item("SHIP_VIA_CODE") = "STANDARD"
                                        End If

                                        If rowSOTORDRX.Item("SHIP_VIA_CODE") & String.Empty = String.Empty Then
                                            rowSOTORDRX.Item("SHIP_VIA_CODE") = (rowData.Item("Id") & String.Empty).ToString.Trim
                                        End If
                                    End If

                                End If

                                ' If not DPD and Not Standard delivery, then lock the ship via
                                If rowSOTORDRX.Item("ORDR_DPD") & String.Empty <> "1" _
                                AndAlso (rowSOTORDRX.Item("SHIP_VIA_CODE") & String.Empty).ToString.Trim.Length > 0 _
                                AndAlso rowSOTORDRX.Item("SHIP_VIA_CODE") & String.Empty <> "STANDARD" Then
                                    rowSOTORDRX.Item("ORDR_LOCK_SHIP_VIA") = "1"
                                End If

                                ' Get DPD Address
                                If DELIVERY_Id.Length > 0 AndAlso rowSOTORDRX.Item("ORDR_DPD") = "1" Then
                                    If vwXmlDataset.Tables("ADDRESS").Select("DELIVERY_Id = " & DELIVERY_Id).Length > 0 Then
                                        rowData = vwXmlDataset.Tables("ADDRESS").Select("DELIVERY_Id = " & DELIVERY_Id)(0)
                                        rowSOTORDRX.Item("CUST_NAME") = TruncateField(AttentionTo, "SOTORDRX", "CUST_NAME")
                                        rowSOTORDRX.Item("CUST_ADDR1") = TruncateField(((rowData.Item("Street_Number") & String.Empty).ToString.Trim & " " & rowData.Item("Street_Name") & String.Empty).ToString.Trim, "SOTORDRX", "CUST_ADDR1")
                                        rowSOTORDRX.Item("CUST_ADDR2") = TruncateField((rowData.Item("Suite") & String.Empty).ToString.Trim, "SOTORDRX", "CUST_ADDR2")
                                        rowSOTORDRX.Item("CUST_CITY") = TruncateField((rowData.Item("City") & String.Empty).ToString.Trim, "SOTORDRX", "CUST_CITY")
                                        rowSOTORDRX.Item("CUST_STATE") = TruncateField((rowData.Item("State") & String.Empty).ToString.Trim, "SOTORDRX", "CUST_STATE")
                                        rowSOTORDRX.Item("CUST_ZIP_CODE") = TruncateField((rowData.Item("Zipcode") & String.Empty).ToString.Trim, "SOTORDRX", "CUST_ZIP_CODE")
                                        Telephone = (rowData.Item("TEL") & String.Empty).ToString.Trim
                                        Telephone = FormatTelePhone(Telephone)
                                        rowSOTORDRX.Item("CUST_PHONE") = TruncateField(Telephone, "SOTORDRX", "CUST_PHONE")
                                        rowSOTORDRX.Item("CUST_FAX") = String.Empty
                                        rowSOTORDRX.Item("CUST_EMAIL") = String.Empty
                                        rowSOTORDRX.Item("CUST_COUNTRY") = TruncateField((rowData.Item("Country") & String.Empty).ToString.Trim, "SOTORDRX", "CUST_COUNTRY")
                                    End If
                                End If

                                SOFTCONTACTS_id = String.Empty
                                SOFTCONTACT_id = String.Empty
                                ORDR_LNO = 1

                                For Each rowSOFTCONTACTS As DataRow In vwXmlDataset.Tables("SOFTCONTACTS").Select(tableName & "_Id = " & SOFT_Id, "SOFTCONTACTS_id")
                                    SOFTCONTACTS_id = rowSOFTCONTACTS.Item("SOFTCONTACTS_id") & String.Empty

                                    For Each rowItem As DataRow In vwXmlDataset.Tables("SOFTCONTACT").Select("SOFTCONTACTS_id = " & SOFTCONTACTS_id, "LineItemID")

                                        SOFTCONTACT_id = rowItem.Item("SOFTCONTACT_id") & String.Empty

                                        ' Need to make header record data for multi item order
                                        If ORDR_LNO > 1 Then
                                            Dim rowHeaderX As DataRow = dst.Tables("SOTORDRX").NewRow
                                            For Each col As DataColumn In dst.Tables("SOTORDRX").Columns
                                                rowHeaderX.Item(col.ColumnName) = rowSOTORDRX.Item(col.ColumnName)
                                            Next
                                            rowHeaderX.Item("ORDR_LNO") = ORDR_LNO
                                            rowHeaderX.Item("ITEM_SPHERE_POWER") = System.DBNull.Value
                                            rowHeaderX.Item("ITEM_ADD_POWER") = System.DBNull.Value
                                            dst.Tables("SOTORDRX").Rows.Add(rowHeaderX)
                                            rowSOTORDRX = rowHeaderX
                                        End If

                                        rowSOTORDRX.Item("ORDR_LNO") = ORDR_LNO
                                        ORDR_LNO += 1
                                        rowSOTORDRX.Item("CUST_LINE_REF") = rowItem.Item("LineItemId") & String.Empty

                                        Dim fieldName As String = IIf(tableName = "STK_SOFT_OFFICE", "SOFTCONTACT_id", "RX_SOFT_PATIENT_ID")
                                        If vwXmlDataset.Tables("PATIENT").Select(fieldName & " = " & SOFTCONTACT_id).Length > 0 Then
                                            rowData = vwXmlDataset.Tables("PATIENT").Select(fieldName & " = " & SOFTCONTACT_id)(0)
                                            rowSOTORDRX.Item("PATIENT_NAME") = TruncateField(((rowData.Item("FirstName") & String.Empty).ToString.Trim & " " & (rowData.Item("LastName") & String.Empty)).ToString.Trim, "SOTORDRX", "PATIENT_NAME")
                                            rowSOTORDRX.Item("PATIENT_NAME") = StrConv(rowSOTORDRX.Item("PATIENT_NAME") & String.Empty, VbStrConv.ProperCase)
                                        End If

                                        ' If it is a DPD and The Patient Name is Blank then use the name on the DPD Address
                                        If (rowSOTORDRX.Item("ORDR_DPD") & String.Empty) = "1" AndAlso (rowSOTORDRX.Item("PATIENT_NAME") & String.Empty).ToString.Trim.Length = 0 Then
                                            rowSOTORDRX.Item("PATIENT_NAME") = StrConv(rowSOTORDRX.Item("CUST_NAME") & String.Empty, VbStrConv.ProperCase)
                                        End If

                                        ' Item Specific
                                        rowSOTORDRX.Item("ORDR_QTY") = Val(rowItem.Item("Quantity") & String.Empty)
                                        rowSOTORDRX.Item("ORDR_UNIT_PRICE_PATIENT") = Val(rowItem.Item("UnitPrice") & String.Empty)
                                        rowSOTORDRX.Item("ORDR_LR") = TruncateField(rowItem.Item("Eye") & String.Empty, "SOTORDRX", "ORDR_LR")
                                        If Not "RL".Contains(rowSOTORDRX.Item("ORDR_LR") & String.Empty) Then
                                            rowSOTORDRX.Item("ORDR_LR") = String.Empty
                                        End If

                                        rowSOTORDRX.Item("ORDR_LINE_SOURCE") = ORDR_SOURCE

                                        itemCode = (rowItem.Item("Id") & String.Empty).ToString.Trim
                                        upcCode = (rowItem.Item("UPCCode") & String.Empty).ToString.Trim

                                        ' As per Maria Use Item Code first.
                                        If itemCode.Length > 0 _
                                            AndAlso ABSolution.ASCDATA1.GetDataRow("Select * From ICTCATL1 Where ITEM_CODE = :PARM1", "V", New Object() {itemCode}) IsNot Nothing Then
                                            rowSOTORDRX.Item("ITEM_CODE") = itemCode
                                        ElseIf upcCode.Length > 0 _
                                            AndAlso ABSolution.ASCDATA1.GetDataRow("Select * From ICTCATL1 Where ITEM_PROD_ID = :PARM1", "V", New Object() {upcCode}) IsNot Nothing Then
                                            rowSOTORDRX.Item("ITEM_CODE") = upcCode
                                        ElseIf itemCode.Length > 0 Then
                                            rowSOTORDRX.Item("ITEM_CODE") = itemCode
                                        Else
                                            rowSOTORDRX.Item("ITEM_CODE") = upcCode
                                        End If

                                        rowSOTORDRX.Item("ITEM_DESC") = TruncateField(rowItem.Item("Name") & String.Empty, "SOTORDRX", "ITEM_DESC")
                                        rowSOTORDRX.Item("ITEM_DESC2") = TruncateField(rowItem.Item("Family") & String.Empty, "SOTORDRX", "ITEM_DESC2")
                                        rowSOTORDRX.Item("ITEM_BASE_CURVE") = Val(rowItem.Item("BaseCurveName") & String.Empty)
                                        rowSOTORDRX.Item("ITEM_DIAMETER") = Val(rowItem.Item("Diameter") & String.Empty)
                                        rowSOTORDRX.Item("ITEM_COLOR") = TruncateField(rowItem.Item("Color") & String.Empty, "SOTORDRX", "ITEM_COLOR")
                                        rowSOTORDRX.Item("ITEM_MULTIFOCAL") = TruncateField(rowItem.Item("GeometricDesignName") & String.Empty, "SOTORDRX", "ITEM_MULTIFOCAL")

                                        If vwXmlDataset.Tables("PRESCRIPTION").Select("SOFTCONTACT_id = " & SOFTCONTACT_id).Length > 0 Then
                                            rowData = vwXmlDataset.Tables("PRESCRIPTION").Select("SOFTCONTACT_id = " & SOFTCONTACT_id)(0)

                                            If rowData.Item("SPHERE") & String.Empty <> String.Empty Then
                                                rowSOTORDRX.Item("ITEM_SPHERE_POWER") = Val(rowData.Item("SPHERE") & String.Empty)
                                            End If

                                            PRESCRIPTION_id = rowData.Item("PRESCRIPTION_id") & String.Empty
                                            If vwXmlDataset.Tables("ADDITION").Select("PRESCRIPTION_id = " & PRESCRIPTION_id).Length > 0 Then
                                                rowData = vwXmlDataset.Tables("ADDITION").Select("PRESCRIPTION_id = " & PRESCRIPTION_id)(0)

                                                If rowData.Item("Value") & String.Empty <> String.Empty Then
                                                    rowSOTORDRX.Item("ITEM_ADD_POWER") = Val(rowData.Item("Value") & String.Empty)
                                                End If
                                            End If

                                            If vwXmlDataset.Tables("CYLINDER").Select("PRESCRIPTION_id = " & PRESCRIPTION_id).Length > 0 Then
                                                rowData = vwXmlDataset.Tables("CYLINDER").Select("PRESCRIPTION_id = " & PRESCRIPTION_id)(0)

                                                If rowData.Item("Axis") & String.Empty <> String.Empty Then
                                                    rowSOTORDRX.Item("ITEM_AXIS") = Val(rowData.Item("Axis") & String.Empty)
                                                End If

                                                If rowData.Item("Value") & String.Empty <> String.Empty Then
                                                    rowSOTORDRX.Item("ITEM_CYLINDER") = Val(rowData.Item("Value") & String.Empty)
                                                End If
                                            End If
                                        End If
                                    Next

                                    ' Not Used here
                                    'rowSOTORDRX.Item("BILLING_NAME") = String.Empty
                                    'rowSOTORDRX.Item("BILLING_ADDRESS1") = String.Empty
                                    'rowSOTORDRX.Item("BILLING_ADDRESS2") = String.Empty
                                    'rowSOTORDRX.Item("BILLING_CITY") = String.Empty
                                    'rowSOTORDRX.Item("BILLING_STATE") = String.Empty
                                    'rowSOTORDRX.Item("BILLING_ZIP") = String.Empty

                                    'rowSOTORDRX.Item("ITEM_CYLINDER") = String.Empty
                                    'rowSOTORDRX.Item("ITEM_AXIS") = String.Empty
                                    'rowSOTORDRX.Item("ITEM_PROD_ID") = String.Empty
                                    'rowSOTORDRX.Item("PRICE_CATGY_CODE") = String.Empty

                                    'rowSOTORDRX.Item("ORDR_SHIP_COMPLETE") = String.Empty
                                    'rowSOTORDRX.Item("ORDR_LOCK_SHIP_VIA") = String.Empty
                                    'rowSOTORDRX.Item("OFFICE_WEBSITE") = String.Empty
                                    'rowSOTORDRX.Item("PROCESS_IND") = String.Empty
                                    'rowSOTORDRX.Item("ITEM_UOM") = String.Empty
                                Next
                            Next
                        Next
                    Next

                    ImportedFiles.Add(orderFile)
                Next

                If dst.Tables("SOTORDRX").Rows.Count > 0 Then
                    ' Commit the data from the Excel file and then archive the file
                    Dim UpdateInProcess As Boolean = False
                    With baseClass
                        Try
                            .BeginTrans()
                            UpdateInProcess = True
                            .clsASCBASE1.Update_Record_TDA("SOTORDRX")
                            .CommitTrans()
                            UpdateInProcess = False

                        Catch ex As Exception
                            If UpdateInProcess Then .Rollback()
                            RecordLogEntry("ProcessBLScan: " & ex.Message)
                        End Try

                    End With

                End If

                ' If DPD then set the LR
                Dim ORDR_CALLER_NAME As String = String.Empty
                Dim ORDR_SHIP_COMPLETE As String = String.Empty
                dst.Tables("SOTORDRX").Rows.Clear()

                ' Need to process each order individually for pricing reasons; therefore
                ' need to move the datat to a temp data table and process each order individually
                For Each headers As DataRow In ABSolution.ASCDATA1.GetDataTable("SELECT DISTINCT ORDR_NO FROM SOTORDRX WHERE PROCESS_IND IS NULL AND ORDR_SOURCE = :PARM1", String.Empty, "V", New Object() {ORDR_SOURCE}).Rows
                    ClearDataSetTables(True)
                    ORDR_NO = headers.Item("ORDR_NO") & String.Empty
                    baseClass.clsASCBASE1.Fill_Records("SOTORDRX", New Object() {ORDR_SOURCE, ORDR_NO})

                    If dst.Tables("SOTORDRX").Rows.Count = 0 Then
                        RecordLogEntry("ProcessVisionWebSalesOrders: Invalid Order Number (" & ORDR_NO & ") for " & vwConnection.ConnectionDescription)
                        Continue For
                    End If

                    'ORDR_NO = String.Empty
                    ORDR_CALLER_NAME = dst.Tables("SOTORDRX").Rows(0).Item("ORDR_CALLER_NAME") & String.Empty
                    ORDR_SHIP_COMPLETE = dst.Tables("SOTORDRX").Rows(0).Item("ORDR_SHIP_COMPLETE") & String.Empty
                    If CreateSalesOrder(ORDR_NO, False, False, ORDR_SOURCE, ORDR_SOURCE, ORDR_CALLER_NAME, False, ORDR_SOURCE) Then
                        dst.Tables("SOTORDR1").Rows(0).Item("ORDR_SHIP_COMPLETE") = ORDR_SHIP_COMPLETE
                        UpdateDataSetTables()
                        salesordersprocessed += 1
                        Try
                            ABSolution.ASCDATA1.ExecuteSQL("DELETE FROM SOTORDRX WHERE ORDR_SOURCE = :PARM1 AND ORDR_NO = :PARM2", _
                                                           "VV", _
                                                        New Object() {ORDR_SOURCE, dst.Tables("SOTORDRX").Rows(0).Item("ORDR_NO") & String.Empty})

                        Catch ex As Exception

                        End Try
                    Else
                        RecordLogEntry("ProcessVisionWebSalesOrders: " & "Could not create sales order for " & dst.Tables("SOTORDRX").Rows(0).Item("ORDR_NO"))
                    End If
                Next
            Catch ex As Exception
                RecordLogEntry("ProcessVisionWebSalesOrders: " & ex.Message)
                emailErrors(ORDR_SOURCE, 1, "Error loading VW file " & fileprocessing & ": " & ex.Message)
            Finally
                RecordLogEntry(salesordersprocessed & " " & vwConnection.ConnectionDescription & " Sales Orders imported.")

                ' Move Xml files to the archive directory
                For Each orderFile As String In ImportedFiles
                    My.Computer.FileSystem.MoveFile(orderFile, vwConnection.LocalInDirArchive & My.Computer.FileSystem.GetName(orderFile), True)
                Next

            End Try

        End Sub

        Private Sub ProcessOptiportSalesOrders(ByVal ORDR_SOURCE As String)

            Dim ftpConnection As New Connection(ORDR_SOURCE)
            Dim salesOrdersProcessed As Integer = 0
            Dim sql As String = String.Empty

            Dim rowSOTORDRX As DataRow = Nothing
            Dim tblICTCATL1 As DataTable = Nothing

            Dim CUST_CODE As String = String.Empty
            Dim XML_DOC_SEQ_NO As String = String.Empty
            Dim ITEM_UPC_CODE As String = String.Empty
            Dim ITEM_DESC2 As String = String.Empty
            Dim ORDR_LNO As Int16 = 0

            Dim ORDR_LINE_SOURCE As String = ORDR_SOURCE
            Dim CREATE_SHIP_TO As Boolean = False
            Dim SELECT_SHIP_TO_BY_TELE As Boolean = False
            Dim CALLER_NAME As String = String.Empty
            Dim ORDR_SHIP_COMPLETE As String = "0"
            Dim XMT_ORDER_SOURCE As String = ORDR_SOURCE

            If testMode Then
                RecordLogEntry("Enter ProcessOptiPort")
            End If

            Try
                sql = "SELECT * From XMTORDR1 Where XML_PROCESS_IND IS NULL"
                baseClass.clsASCBASE1.Fill_Records("XMTORDR1", String.Empty, False, sql)

                If dst.Tables("XMTXREF1").Select("ORDR_LINE_SOURCE = '" & ORDR_SOURCE & "'", "").Length > 0 Then
                    With dst.Tables("XMTXREF1").Select("ORDR_LINE_SOURCE = '" & ORDR_SOURCE & "'", "")(0)
                        ORDR_LINE_SOURCE = .Item("ORDR_LINE_SOURCE") & String.Empty
                        XMT_ORDER_SOURCE = .Item("ORDR_SOURCE") & String.Empty
                        CREATE_SHIP_TO = (.Item("CREATE_SHIP_TO") & String.Empty) = "1"
                        SELECT_SHIP_TO_BY_TELE = (.Item("SELECT_SHIP_TO_BY_TELE") & String.Empty) = "1"
                        CALLER_NAME = .Item("CALLER_NAME") & String.Empty
                        ORDR_SHIP_COMPLETE = .Item("ORDR_SHIP_COMPLETE") & String.Empty
                        If ORDR_SHIP_COMPLETE.Length = 0 Then ORDR_SHIP_COMPLETE = "0"
                    End With
                End If

                For Each rowXMTORDR1 As DataRow In dst.Tables("XMTORDR1").Select("", "XML_DOC_SEQ_NO")
                    ClearDataSetTables(False)

                    ' Done so if the record causes an error on the import it will be skipped.
                    ' An email with the error will be sent.
                    rowXMTORDR1.Item("XML_PROCESS_IND") = "X"
                    With baseClass
                        Try
                            .BeginTrans()
                            .clsASCBASE1.Update_Record_TDA("XMTORDR1")
                            .CommitTrans()
                        Catch ex As Exception
                            ' nothing at this time
                        End Try
                    End With

                    rowXMTORDR1.Item("XML_PROCESS_IND") = "1"
                    XML_DOC_SEQ_NO = rowXMTORDR1.Item("XML_DOC_SEQ_NO") & String.Empty
                    CUST_CODE = (rowXMTORDR1.Item("XML_CUSTOMER_ID") & String.Empty).ToString.ToUpper.Trim
                    CUST_CODE = ABSolution.ASCMAIN1.Format_Field(CUST_CODE, "CUST_CODE")
                    ORDR_LNO = 1

                    baseClass.clsASCBASE1.Fill_Records("XMTORDR2", XML_DOC_SEQ_NO)

                    For Each rowXMTORDR2 As DataRow In dst.Tables("XMTORDR2").Select("", "XML_DOC_SEQ_NO")

                        rowSOTORDRX = dst.Tables("SOTORDRX").NewRow
                        rowSOTORDRX.Item("ORDR_SOURCE") = XMT_ORDER_SOURCE
                        rowSOTORDRX.Item("ORDR_NO") = rowXMTORDR1.Item("XML_ORDER_ID") & String.Empty
                        rowSOTORDRX.Item("ORDR_LNO") = ORDR_LNO
                        rowSOTORDRX.Item("ORDR_CUST_PO") = rowXMTORDR1.Item("XML_PO_NO") & String.Empty
                        rowSOTORDRX.Item("CUST_CODE") = CUST_CODE
                        rowSOTORDRX.Item("CUST_SHIP_TO_NO") = String.Empty
                        rowSOTORDRX.Item("ORDR_DATE") = Format(DateTime.Now, "MM/dd/yyyy")
                        rowSOTORDRX.Item("ORDR_SHIP_COMPLETE") = ORDR_SHIP_COMPLETE
                        ' needed for telephone look up
                        rowSOTORDRX.Item("CUST_SHIP_TO_PHONE") = (rowXMTORDR1.Item("XML_OFFICE_TEL") & String.Empty).ToString.Replace(" ", "")

                        rowSOTORDRX.Item("PATIENT_NAME") = rowXMTORDR2.Item("XML_PATIENT_NAME") & String.Empty
                        rowSOTORDRX.Item("PATIENT_NAME") = TruncateField(rowSOTORDRX.Item("PATIENT_NAME") & String.Empty, "SOTORDR2", "PATIENT_NAME")
                        rowSOTORDRX.Item("PATIENT_NAME") = StrConv(rowSOTORDRX.Item("PATIENT_NAME") & String.Empty, VbStrConv.ProperCase)

                        If rowXMTORDR1.Item("XML_SHIP_TO_PATIENT") & String.Empty = "Y" Then
                            rowSOTORDRX.Item("ORDR_DPD") = "1"
                            rowSOTORDRX.Item("ORDR_SHIP_COMPLETE") = "1"

                            rowSOTORDRX.Item("CUST_NAME") = (rowXMTORDR1.Item("XML_SHIP_TO_NAME") & String.Empty)
                            rowSOTORDRX.Item("CUST_ADDR1") = (rowXMTORDR1.Item("XML_SHIP_TO_ADDRESS1") & String.Empty).ToString.Trim
                            rowSOTORDRX.Item("CUST_ADDR2") = (rowXMTORDR1.Item("XML_SHIP_TO_ADDRESS2") & String.Empty).ToString.Trim
                            rowSOTORDRX.Item("CUST_CITY") = (rowXMTORDR1.Item("XML_SHIP_TO_CITY") & String.Empty).ToString.Trim
                            rowSOTORDRX.Item("CUST_STATE") = (rowXMTORDR1.Item("XML_SHIP_TO_STATE") & String.Empty).ToString.Trim
                            rowSOTORDRX.Item("CUST_ZIP_CODE") = (rowXMTORDR1.Item("XML_SHIP_TO_ZIP") & String.Empty).ToString.Trim
                            rowSOTORDRX.Item("CUST_COUNTRY") = "US"
                        Else
                            rowSOTORDRX.Item("ORDR_DPD") = "0"
                        End If

                        rowSOTORDRX.Item("SHIP_VIA_CODE") = (rowXMTORDR1.Item("XML_SHIPPING_METHOD") & String.Empty).ToString.ToUpper.Trim

                        rowSOTORDRX.Item("EDI_CUST_REF_NO") = (rowXMTORDR1.Item("XML_ORDER_ID") & String.Empty).ToString.Trim
                        rowSOTORDRX.Item("ORDR_TYPE_CODE") = "REG"
                        rowSOTORDRX.Item("ORDR_CALLER_NAME") = CALLER_NAME

                        rowSOTORDRX.Item("CUST_LINE_REF") = rowXMTORDR2.Item("XML_ITEM_ID") & String.Empty
                        rowSOTORDRX.Item("ORDR_QTY") = Val(rowXMTORDR2.Item("XML_ORDER_QTY") & String.Empty)
                        rowSOTORDRX.Item("ORDR_UNIT_PRICE_PATIENT") = 0
                        Select Case rowXMTORDR2.Item("XML_ITEM_EYE") & String.Empty
                            Case "OD"
                                rowSOTORDRX.Item("ORDR_LR") = "R"
                            Case "OS"
                                rowSOTORDRX.Item("ORDR_LR") = "L"
                        End Select

                        rowSOTORDRX.Item("ORDR_LINE_SOURCE") = ORDR_LINE_SOURCE

                        ITEM_UPC_CODE = (rowXMTORDR2.Item("XML_UPC_CODE") & String.Empty).ToString.Trim
                        If ITEM_UPC_CODE.Trim.Length > 0 Then
                            ITEM_UPC_CODE = ABSolution.ASCMAIN1.Format_Field(ITEM_UPC_CODE, "ITEM_UPC_CODE")
                            rowSOTORDRX.Item("ITEM_CODE") = ITEM_UPC_CODE
                        Else
                            sql = "  SELECT * FROM ICTCATL1"
                            sql = sql & " WHERE  PRICE_CATGY_CODE = :PARM1"
                            sql = sql & " AND ITEM_BASE_CURVE = :PARM2"
                            sql = sql & " AND ITEM_DIAMETER = :PARM3"
                            sql = sql & " AND ITEM_SPHERE_POWER = :PARM4"
                            sql = sql & " AND ITEM_CYLINDER = :PARM5"
                            sql = sql & " AND ITEM_AXIS = :PARM6"
                            sql = sql & " AND ITEM_ADD_POWER = :PARM7"
                            sql = sql & " AND ITEM_COLOR = :PARM8"
                            sql = sql & " AND ITEM_ADD_DOM_NON = :PARM9"
                            tblICTCATL1 = ABSolution.ASCDATA1.GetDataTable(sql, "ICTCATL1", "VVVVVVVVV", _
                                        New Object() {rowXMTORDR2.Item("XML_PRODUCT_KEY") & String.Empty, _
                                                       rowXMTORDR2.Item("XML_ITEM_BASE_CURVE") & String.Empty, _
                                                       rowXMTORDR2.Item("XML_ITEM_DIAMETER") & String.Empty, _
                                                       rowXMTORDR2.Item("XML_ITEM_SPHERE_POWER") & String.Empty, _
                                                       rowXMTORDR2.Item("XML_ITEM_CYLINDER") & String.Empty, _
                                                       rowXMTORDR2.Item("XML_ITEM_AXIS") & String.Empty, _
                                                       rowXMTORDR2.Item("XML_ITEM_ADD_POWER") & String.Empty, _
                                                       rowXMTORDR2.Item("XML_ITEM_COLOR") & String.Empty, _
                                                       rowXMTORDR2.Item("XML_ITEM_MULTIFOCAL") & String.Empty})

                            If tblICTCATL1.Rows.Count > 0 Then
                                ITEM_UPC_CODE = tblICTCATL1.Rows(0).Item("ITEM_PROD_ID") & String.Empty
                            End If

                            rowSOTORDRX.Item("PRICE_CATGY_CODE") = rowXMTORDR2.Item("XML_PRODUCT_KEY") & String.Empty
                            rowSOTORDRX.Item("ITEM_BASE_CURVE") = Val(rowXMTORDR2.Item("XML_ITEM_BASE_CURVE") & String.Empty)
                            rowSOTORDRX.Item("ITEM_DIAMETER") = Val(rowXMTORDR2.Item("XML_ITEM_DIAMETER") & String.Empty)
                            rowSOTORDRX.Item("ITEM_SPHERE_POWER") = Val(rowXMTORDR2.Item("XML_ITEM_SPHERE_POWER") & String.Empty)
                            rowSOTORDRX.Item("ITEM_CYLINDER") = Val(rowXMTORDR2.Item("XML_ITEM_CYLINDER") & String.Empty)
                            rowSOTORDRX.Item("ITEM_AXIS") = Val(rowXMTORDR2.Item("XML_ITEM_AXIS") & String.Empty)
                            rowSOTORDRX.Item("ITEM_ADD_POWER") = Val(rowXMTORDR2.Item("XML_ITEM_ADD_POWER") & String.Empty)

                            rowSOTORDRX.Item("ITEM_COLOR") = TruncateField(rowXMTORDR2.Item("XML_ITEM_COLOR") & String.Empty, "SOTORDRX", "ITEM_COLOR")
                            rowSOTORDRX.Item("ITEM_MULTIFOCAL") = TruncateField(rowXMTORDR2.Item("XML_ITEM_MULTIFOCAL") & String.Empty, "SOTORDRX", "ITEM_MULTIFOCAL")


                            ITEM_DESC2 = rowXMTORDR2.Item("XML_PRODUCT_KEY") & String.Empty & "/" & _
                               rowXMTORDR2.Item("XML_ITEM_BASE_CURVE") & String.Empty & "/" & _
                               rowXMTORDR2.Item("XML_ITEM_DIAMETER") & String.Empty & "/" & _
                               rowXMTORDR2.Item("XML_ITEM_SPHERE_POWER") & String.Empty & "/" & _
                               rowXMTORDR2.Item("XML_ITEM_CYLINDER") & String.Empty & "/" & _
                               rowXMTORDR2.Item("XML_ITEM_AXIS") & String.Empty & "/" & _
                               rowXMTORDR2.Item("XML_ITEM_ADD_POWER") & String.Empty & "/" & _
                               rowXMTORDR2.Item("XML_ITEM_COLOR") & String.Empty & "/" & _
                               rowXMTORDR2.Item("XML_ITEM_MULTIFOCAL") & String.Empty

                            ITEM_DESC2 = ITEM_DESC2.Trim
                            If ITEM_DESC2.Length > 60 Then
                                ITEM_DESC2 = ITEM_DESC2.Substring(0, 60).Trim
                            End If
                            rowSOTORDRX.Item("ITEM_DESC2") = ITEM_DESC2
                        End If

                        rowSOTORDRX.Item("ITEM_DESC") = String.Empty
                        rowSOTORDRX.Item("ITEM_PROD_ID") = ITEM_UPC_CODE
                        rowSOTORDRX.Item("ITEM_CODE") = ITEM_UPC_CODE

                        rowSOTORDRX.Item("BILLING_NAME") = String.Empty
                        rowSOTORDRX.Item("BILLING_ADDRESS1") = String.Empty
                        rowSOTORDRX.Item("BILLING_ADDRESS2") = String.Empty
                        rowSOTORDRX.Item("BILLING_CITY") = String.Empty
                        rowSOTORDRX.Item("BILLING_STATE") = String.Empty
                        rowSOTORDRX.Item("BILLING_ZIP") = String.Empty

                        rowSOTORDRX.Item("ORDR_LOCK_SHIP_VIA") = String.Empty
                        rowSOTORDRX.Item("ORDR_COMMENT") = String.Empty
                        rowSOTORDRX.Item("OFFICE_WEBSITE") = String.Empty
                        rowSOTORDRX.Item("PROCESS_IND") = String.Empty
                        rowSOTORDRX.Item("ITEM_UOM") = String.Empty
                        dst.Tables("SOTORDRX").Rows.Add(rowSOTORDRX)

                        ORDR_LNO += 1
                    Next

                    ' Create sales order
                    Dim ORDR_NO As String = String.Empty
                    If CreateSalesOrder(ORDR_NO, CREATE_SHIP_TO, SELECT_SHIP_TO_BY_TELE, ORDR_LINE_SOURCE, XMT_ORDER_SOURCE, CALLER_NAME, False, ORDR_SOURCE) Then
                        UpdateDataSetTables()
                        salesOrdersProcessed += 1
                    Else
                        rowXMTORDR1.Item("XML_PROCESS_IND") = "B"
                        RecordLogEntry("OptiPort XML_DOC_SEQ_NO : " & (rowXMTORDR1.Item("XML_DOC_SEQ_NO") & String.Empty).ToString & " not imported")
                        emailErrors("X", 1, "OptiPort XML_DOC_SEQ_NO: " & (rowXMTORDR1.Item("XML_DOC_SEQ_NO") & String.Empty).ToString & " not imported")

                        With baseClass
                            Try
                                .BeginTrans()
                                .clsASCBASE1.Update_Record_TDA("XMTORDR1")
                                .CommitTrans()
                            Catch ex As Exception
                                .Rollback()
                                RecordLogEntry("UpdateDataSetTables  : " & ex.Message)
                            End Try
                        End With
                    End If
                Next

                If testMode Then RecordLogEntry("Exit ProcessWebServiceSalesOrders.")

            Catch ex As Exception
                RecordLogEntry("ProcessOptiPort: " & ex.Message)
                emailErrors("X", 1, ex.Message)
            Finally
                RecordLogEntry(salesOrdersProcessed & " " & ftpConnection.ConnectionDescription & " Sales Orders imported.")
                If testMode Then
                    RecordLogEntry("End ProcessOptiPort")
                End If
            End Try
        End Sub

        Private Sub ExportVisionWebStatus()

            Dim numOrdersProcessed As Int16 = 0
            Dim vwConnection As New Connection("V")
            Dim xmlFilename As String = String.Empty
            Dim xmlWriter As XmlTextWriter

            Dim statusDesc As List(Of String)
            Dim ordersProcessed As List(Of String) = New List(Of String)

            Dim ORDR_NO As String = String.Empty
            Dim sql As String = String.Empty

            Dim tblSotinvh1 As DataTable = Nothing
            Dim rowSotinvh1 As DataRow = Nothing

            Dim ORDR_QTY_OPEN As Integer = 0
            Dim ORDR_QTY_PICK As Integer = 0
            Dim ORDR_QTY_SHIP As Integer = 0
            Dim ORDR_QTY_CANC As Integer = 0
            Dim ORDR_QTY_BACK As Integer = 0
            Dim ORDR_QTY_ONPO As Integer = 0
            Dim ORDR_LINE_STATUS As String = String.Empty
            Dim VWebStatus As String = String.Empty

            Try
                Dim rowSOTORDR1 As DataRow = Nothing
                Dim tblSOTORDR2 As DataTable = Nothing

                baseClass.clsASCBASE1.Fill_Records("XSTORDRQ", String.Empty, True, "Select * From XSTORDRQ WHERE ORDR_SOURCE = 'V'")

                For Each rowXSTORDRQ As DataRow In dst.Tables("XSTORDRQ").Select("", "ORDR_NO, LAST_DATE DESC")

                    ORDR_NO = rowXSTORDRQ.Item("ORDR_NO")

                    If ordersProcessed.Contains(ORDR_NO) Then
                        Continue For
                    End If

                    ordersProcessed.Add(ORDR_NO)

                    rowSOTORDR1 = ABSolution.ASCDATA1.GetDataRow("Select * From SOTORDR1 Where ORDR_NO = :PARM1", "V", New Object() {ORDR_NO})
                    tblSOTORDR2 = ABSolution.ASCDATA1.GetDataTable("Select * From SOTORDR2 Where ORDR_NO = :PARM1", "", "V", New Object() {ORDR_NO})

                    ' Send over the Shipping / Tracking / Url for the shipment
                    sql = " SELECT SOTINVH1.INV_NO, SOTINVH2.INV_LNO, SOTINVH1.INV_DATE, SOTINVH1.SHIP_VIA_CODE, SOTSVIA1.SHIP_VIA_DESC, SOTINVH1.SHIP_REF"
                    sql &= " FROM SOTINVH1, SOTINVH2, SOTSVIA1"
                    sql &= " WHERE SOTINVH1.INV_NO = SOTINVH2.INV_NO"
                    sql &= " AND SOTINVH1.INV_TYPE = SOTINVH1.INV_TYPE"
                    sql &= " AND SOTSVIA1.SHIP_VIA_CODE = SOTINVH1.SHIP_VIA_CODE"
                    sql &= " AND SOTINVH1.ORDR_NO = :PARM1"

                    tblSotinvh1 = ABSolution.ASCDATA1.GetDataTable(sql, "", "V", New Object() {ORDR_NO})

                    ' We need to have the Vision Web ID code.
                    ' The Import uses EDI_CUST_REF_NO and keyed in orders use ORDR_CUST_PO
                    If rowSOTORDR1.Item("EDI_CUST_REF_NO") & String.Empty = String.Empty AndAlso _
                        rowSOTORDR1.Item("ORDR_CUST_PO") & String.Empty = String.Empty Then
                        Continue For
                    End If

                    numOrdersProcessed += 1

                    xmlFilename = ORDR_NO & DateTime.Now.ToString("_yyyyMMddhhmmss") & ".xml"

                    If My.Computer.FileSystem.FileExists(vwConnection.LocalOutDir & xmlFilename) Then
                        My.Computer.FileSystem.DeleteFile(vwConnection.LocalOutDir & xmlFilename)
                    End If

                    xmlWriter = New XmlTextWriter(vwConnection.LocalOutDir & xmlFilename, System.Text.Encoding.UTF8)
                    xmlWriter.WriteStartDocument(True)
                    xmlWriter.Formatting = Formatting.Indented
                    xmlWriter.Indentation = 4
                    xmlWriter.WriteStartElement("VW_TRACKING")

                    xmlWriter.WriteStartElement("SUPPLIER")
                    xmlWriter.WriteStartAttribute("Id")
                    xmlWriter.WriteValue("5540")
                    xmlWriter.WriteEndAttribute()

                    xmlWriter.WriteStartElement("ACCOUNT")

                    For Each rowSOTORDR2 As DataRow In tblSOTORDR2.Select("", "ORDR_LNO")
                        xmlWriter.WriteStartElement("ITEM")

                        xmlWriter.WriteStartAttribute("Tracking_Id")
                        xmlWriter.WriteValue(rowSOTORDR1.Item("ORDR_NO") & String.Empty)
                        xmlWriter.WriteEndAttribute()

                        xmlWriter.WriteStartAttribute("Type")
                        If rowSOTORDR1.Item("ORDR_NO") & String.Empty = "1" Then
                            xmlWriter.WriteValue("CP")
                        Else
                            xmlWriter.WriteValue("CO")
                        End If
                        xmlWriter.WriteEndAttribute()

                        xmlWriter.WriteStartAttribute("Visionweb_Tracking_Id")
                        If rowSOTORDR1.Item("EDI_CUST_REF_NO") & String.Empty <> String.Empty Then
                            xmlWriter.WriteValue(rowSOTORDR1.Item("EDI_CUST_REF_NO") & String.Empty)
                        ElseIf rowSOTORDR1.Item("ORDR_CUST_PO") & String.Empty <> String.Empty Then
                            xmlWriter.WriteValue((rowSOTORDR1.Item("ORDR_CUST_PO") & String.Empty).ToString.ToUpper.Trim)
                        End If
                        xmlWriter.WriteEndAttribute()

                        xmlWriter.WriteStartAttribute("Patient")
                        xmlWriter.WriteValue(rowSOTORDR2.Item("PATIENT_NAME") & String.Empty)
                        xmlWriter.WriteEndAttribute()

                        xmlWriter.WriteStartAttribute("Received_at")
                        xmlWriter.WriteValue(rowSOTORDR1.Item("INIT_DATE") & String.Empty)
                        xmlWriter.WriteEndAttribute()

                        xmlWriter.WriteStartAttribute("Origin")
                        If rowSOTORDR1.Item("EDI_CUST_REF_NO") & String.Empty = String.Empty Then
                            xmlWriter.WriteValue("PHN")
                        Else
                            xmlWriter.WriteValue("VWB")
                        End If
                        xmlWriter.WriteEndAttribute()

                        statusDesc = New List(Of String)

                        ORDR_QTY_OPEN = Val(rowSOTORDR2.Item("ORDR_QTY_OPEN") & String.Empty)
                        If ORDR_QTY_OPEN > 0 Then
                            statusDesc.Add(rowSOTORDR2.Item("ITEM_CODE") & ": " & ORDR_QTY_OPEN & " piece(s) Open")
                        End If

                        ORDR_QTY_PICK = Val(rowSOTORDR2.Item("ORDR_QTY_PICK") & String.Empty)
                        If ORDR_QTY_PICK > 0 Then
                            statusDesc.Add(rowSOTORDR2.Item("ITEM_CODE") & ": " & ORDR_QTY_PICK & " piece(s) sent to warehouse")
                        End If

                        ORDR_QTY_SHIP = Val(rowSOTORDR2.Item("ORDR_QTY_SHIP") & String.Empty)
                        If ORDR_QTY_SHIP > 0 Then
                            statusDesc.Add(rowSOTORDR2.Item("ITEM_CODE") & ": " & ORDR_QTY_SHIP & " piece(s) shipped")
                        End If

                        ORDR_QTY_CANC = Val(rowSOTORDR2.Item("ORDR_QTY_CANC") & String.Empty)
                        If ORDR_QTY_CANC > 0 Then
                            statusDesc.Add(rowSOTORDR2.Item("ITEM_CODE") & ": " & ORDR_QTY_CANC & " piece(s) cancelled")
                        End If

                        ORDR_QTY_BACK = Val(rowSOTORDR2.Item("ORDR_QTY_BACK") & String.Empty)
                        If ORDR_QTY_BACK > 0 Then
                            statusDesc.Add(rowSOTORDR2.Item("ITEM_CODE") & ": " & ORDR_QTY_BACK & " piece(s) backordered")
                        End If

                        ORDR_QTY_ONPO = Val(rowSOTORDR2.Item("ORDR_QTY_ONPO") & String.Empty)
                        If ORDR_QTY_ONPO > 0 Then
                            statusDesc.Add(rowSOTORDR2.Item("ITEM_CODE") & ": " & ORDR_QTY_ONPO & " piece(s) ordered from vendor")
                        End If

                        ORDR_LINE_STATUS = rowSOTORDR2.Item("ORDR_LINE_STATUS") & String.Empty
                        Select Case ORDR_LINE_STATUS
                            Case "F"
                                VWebStatus = CInt(VisionWebStatus.Shipped).ToString.Trim
                            Case "P"
                                VWebStatus = CInt(VisionWebStatus.Shipping).ToString.Trim
                            Case "B"
                                VWebStatus = CInt(VisionWebStatus.Other).ToString.Trim
                            Case "V", "C"
                                VWebStatus = CInt(VisionWebStatus.Cancelled).ToString.Trim
                            Case "O"
                                VWebStatus = CInt(VisionWebStatus.OrderInProcess).ToString.Trim
                            Case Else
                                VWebStatus = CInt(VisionWebStatus.Other).ToString.Trim
                        End Select

                        If rowSOTORDR1.Item("ORDR_HOLD_SALES") & String.Empty = "1" OrElse rowSOTORDR1.Item("ORDR_HOLD_CREDIT") & String.Empty = "1" Then
                            statusDesc.Add("Sales Order on hold")
                        End If

                        xmlWriter.WriteStartAttribute("Status")
                        xmlWriter.WriteValue(VWebStatus)
                        xmlWriter.WriteEndAttribute()

                        ' Write out Status Descriptions
                        Dim desc As String = String.Empty
                        For iLoop As Integer = 0 To statusDesc.Count - 1
                            desc = statusDesc(iLoop)
                            If desc.Length > 59 Then
                                desc = desc.Substring(0, 59).Trim
                            End If
                            xmlWriter.WriteElementString("STATUS_DESCRIPTION", desc)
                        Next

                        xmlWriter.WriteElementString("DESCRIPTION", rowSOTORDR2.Item("ITEM_DESC") & String.Empty)

                        sql = "INV_LNO = " & rowSOTORDR2.Item("ORDR_LNO")
                        If tblSotinvh1.Select(sql).Length > 0 Then
                            rowSotinvh1 = tblSotinvh1.Select(sql, "INV_NO DESC")(0)
                            ' If the invoice is more than 1 day old then do not send shipping info 
                            If Math.Abs(DateDiff(DateInterval.Day, DateTime.Now, CDate(rowSotinvh1.Item("INV_DATE")))) <= 1 Then
                                xmlWriter.WriteStartElement("SHIPPING")

                                xmlWriter.WriteStartAttribute("Tracking")
                                xmlWriter.WriteValue(rowSotinvh1.Item("SHIP_VIA_DESC") & String.Empty)
                                xmlWriter.WriteEndAttribute()

                                xmlWriter.WriteStartAttribute("Url")
                                xmlWriter.WriteValue(Track_Shipment(rowSotinvh1.Item("SHIP_VIA_CODE") & String.Empty, rowSotinvh1.Item("SHIP_REF") & String.Empty))
                                xmlWriter.WriteEndAttribute()

                                xmlWriter.WriteEndElement() ' SHIPPING
                            End If
                        End If
                        xmlWriter.WriteEndElement() ' ITEM
                    Next

                    xmlWriter.WriteEndElement() ' ACCOUNT
                    xmlWriter.WriteEndElement() ' SUPPLIER
                    xmlWriter.WriteEndElement() ' VW_TRACKING
                    xmlWriter.WriteEndDocument()
                    xmlWriter.Close()

                Next

            Catch ex As Exception
                RecordLogEntry("ExportVisionWebStatus: " & ex.Message)
            Finally

                ' Delete processed Orders
                For Each orderNumber As String In ordersProcessed
                    For Each row As DataRow In baseClass.clsASCBASE1.dst.Tables("XSTORDRQ").Select("ORDR_NO = '" & orderNumber & "'")
                        sql = "Delete From XSTORDRQ Where ORDR_NO = '" & orderNumber & "' AND LAST_DATE <= TO_TIMESTAMP('" & row.Item("LAST_DATE") & "','MM/DD/YYYY HH:MI:SS AM') +.00001"
                        ABSolution.ASCDATA1.ExecuteSQL(sql)
                    Next
                Next

                RecordLogEntry("ExportVisionWebStatus: " & numOrdersProcessed & " sales order updates placed in " & vwConnection.LocalOutDir)
            End Try

        End Sub

        Private Function Track_Shipment(ByVal SHIP_VIA_CODE As String, ByVal SHIP_REF As String) As String

            Try

                SHIP_VIA_CODE = SHIP_VIA_CODE.Trim
                SHIP_REF = SHIP_REF.Trim

                If SHIP_VIA_CODE.Length = 0 OrElse SHIP_REF.Length = 0 Then
                    Return String.Empty
                End If

                Dim sql As String = String.Empty
                sql = "Select CARRIER_URL_TRACKING, CARRIER_TRACKING_IND" _
                & " from SOTCARR1,SOTROUT1,SOTSVIA1 " _
                & " where SOTSVIA1.SHIP_VIA_CODE = '" & SHIP_VIA_CODE & "'" _
                & "   and SOTROUT1.ROUTE_CODE = SOTSVIA1.ROUTE_CODE " _
                & "   and SOTCARR1.CARRIER_CODE = SOTROUT1.CARRIER_CODE"

                Dim rowSOTCARR1 As DataRow = ABSolution.ASCDATA1.GetDataRow(sql, True)
                Dim CARRIER_URL_TRACKING As String = rowSOTCARR1.Item("CARRIER_URL_TRACKING") & String.Empty
                Dim CARRIER_TRACKING_IND As String = rowSOTCARR1.Item("CARRIER_TRACKING_IND") & String.Empty

                If CARRIER_TRACKING_IND = "I" Then
                    sql = "SELECT NVL(INV_NO_RESHIP, INV_NO) FROM SOTINVH1 WHERE SHIP_REF = :PARM1 AND SHIP_VIA_CODE = :PARM2"
                    SHIP_REF = ABSolution.ASCDATA1.GetDataValue(sql, "VV", New String() {SHIP_REF, SHIP_VIA_CODE}) & String.Empty
                End If

                If CARRIER_URL_TRACKING = "" Then
                    Return String.Empty
                ElseIf SHIP_REF.Length = 0 AndAlso CARRIER_TRACKING_IND = "I" Then
                    Return String.Empty
                Else
                    Return CARRIER_URL_TRACKING & SHIP_REF
                End If

            Catch ex As Exception
                Return String.Empty
            End Try

        End Function

        Private Sub DisposeOPD()
            Try
                With baseClass.clsASCBASE1

                    If .CMDs IsNot Nothing AndAlso .CMDs.Count <> 0 Then
                        For Each CMD_key As String In .CMDs.Keys
                            Dim cmd As Oracle.DataAccess.Client.OracleCommand = .CMDs(CMD_key)
                            For Each param As Oracle.DataAccess.Client.OracleParameter In cmd.Parameters
                                param.Dispose()
                            Next
                            cmd.Dispose()
                        Next
                    End If
                    .CMDs = Nothing

                    If .BA_CMDs IsNot Nothing AndAlso .BA_CMDs.Count <> 0 Then
                        For Each CMD_key As String In .BA_CMDs.Keys
                            Dim cmds() As Oracle.DataAccess.Client.OracleCommand = .BA_CMDs(CMD_key)
                            For Each cmd As Oracle.DataAccess.Client.OracleCommand In cmds
                                For Each param As Oracle.DataAccess.Client.OracleParameter In cmd.Parameters
                                    param.Dispose()
                                Next
                                cmd.Dispose()
                            Next
                            cmds = Nothing
                        Next
                    End If
                    .BA_CMDs = Nothing

                    If .TDAs IsNot Nothing Then
                        For Each tda As Oracle.DataAccess.Client.OracleDataAdapter In .TDAs.Values
                            tda.Dispose()
                        Next
                    End If
                    .TDAs = Nothing

                    .Dispose()
                End With

                baseClass.Dispose()

            Catch ex As Exception

            End Try
        End Sub

#End Region

#Region "Vision Web DEL Processing"

        Private Sub ProcessVisionWebDELOrders(ByVal ORDR_SOURCE As String)

            ' Hard D to get separate parameters for Vision Web Digital Eyelab orders
            Dim vwConnection As New Connection("D")
            Dim numJobsProcessed As Int16 = 0
            Dim ImportedFiles As List(Of String) = New List(Of String)
            Dim JOB_NO As String = String.Empty
            Dim ORDR_NO As String = String.Empty
            Dim creationDate As String = String.Empty
            Dim rowSOTORDR5 As DataRow = Nothing
            Dim rowData As DataRow = Nothing
            Dim tempString As String = String.Empty

            Dim CUST_CODE As String = String.Empty
            Dim CUST_SHIP_TO_NO As String = String.Empty
            Dim Telephone As String = String.Empty

            Dim RX_SPECTACLE_Id As String = String.Empty
            Dim HEADER_Id As String = String.Empty
            Dim CUSTOMER_Id As String = String.Empty
            Dim ACCOUNTS_Id As String = String.Empty
            Dim ACCOUNT_Id As String = String.Empty
            Dim SP_EQUIPMENT_Id As String = String.Empty
            Dim PATIENT_ID As String = String.Empty
            Dim POSITION_ID As String = String.Empty
            Dim PRESCRIPTION_ID As String = String.Empty
            Dim LENS_ID As String = String.Empty
            Dim TREATMENTS_ID As String = String.Empty
            Dim code As String = String.Empty

            Dim ErrorCodes As List(Of String) = New List(Of String)
            Dim errors As List(Of DelJobService.DelJobValidationError) = New List(Of DelJobService.DelJobValidationError)
            Dim ERR_LNO As Int16 = 0
            Dim inTrans As Boolean = False
            Dim specialInstructions As String = String.Empty
            Dim errorsDirectory As String = String.Empty
            Dim invalidCustomer As Boolean = False

            Dim sqlColorCode As String = String.Empty
            sqlColorCode = " SELECT * FROM DETCOLR1 WHERE COLOR_CODE = "
            sqlColorCode &= " ("
            sqlColorCode &= " SELECT COLOR_CODE FROM (SELECT CX.COLOR_CODE,CASE WHEN NVL(M2.COLOR_CODE_DEFAULT,'')=CX.COLOR_CODE THEN 1 ELSE 0 END SORTRANK"
            sqlColorCode &= " FROM"
            sqlColorCode &= " DETDSGN3 D3"
            sqlColorCode &= " JOIN"
            sqlColorCode &= " (SELECT C1.COLOR_CODE,C2.COLOR_CODE_WEB FROM DETCOLR1 C1 JOIN DETCOLR2 C2 ON (C1.COLOR_CODE_WEB = C2.COLOR_CODE_WEB)) CX"
            sqlColorCode &= "  ON (D3.COLOR_CODE=CX.COLOR_CODE)"
            sqlColorCode &= "  LEFT JOIN"
            sqlColorCode &= "  DETMATL2 M2 ON (M2.MATL_CODE=D3.MATL_CODE AND M2.COLOR_CODE_WEB=CX.COLOR_CODE_WEB)"
            sqlColorCode &= "  WHERE"
            sqlColorCode &= "  D3.LENS_DESIGN_CODE = :PARM1"
            sqlColorCode &= "  AND D3.MATL_CODE = :PARM2"
            sqlColorCode &= "  AND CX.COLOR_CODE_WEB = :PARM3 AND ROWNUM <= 1 ORDER BY SORTRANK DESC) X"
            sqlColorCode &= " )"

            Try
                ImportedFiles.Clear()
                errors.Clear()
                ERR_LNO = 0

                errorsDirectory = vwConnection.LocalInDirArchive
                If Not errorsDirectory.EndsWith("\") Then
                    errorsDirectory &= "\"
                End If
                errorsDirectory &= "Errors\"
                If Not My.Computer.FileSystem.DirectoryExists(errorsDirectory) Then
                    My.Computer.FileSystem.CreateDirectory(errorsDirectory)
                End If

                Dim xsdFile As String = vwConnection.LocalInDir & "CurrentRxOrder.xsd"

                If Not My.Computer.FileSystem.FileExists(xsdFile) Then
                    RecordLogEntry("ProcessVisionWebDELOrders: " & xsdFile & " could not be found.")
                    emailErrors(String.Empty, 1, "ProcessVisionWebDELOrders: " & xsdFile & " could not be found.", "DEL_VWEB")
                    Exit Sub
                End If

                ' Create the DataSet to read the schema into.
                vwXmlDataset = New DataSet
                'Create a FileStream object with the file path and name.
                Dim myFileStream As System.IO.FileStream = New System.IO.FileStream(xsdFile, System.IO.FileMode.Open)
                'Create a new XmlTextReader object with the FileStream.
                Dim myXmlTextReader As System.Xml.XmlTextReader = New System.Xml.XmlTextReader(myFileStream)
                'Read the schema into the DataSet and close the reader.
                vwXmlDataset.ReadXmlSchema(myXmlTextReader)
                myXmlTextReader.Close()

                ' Loop through downloaded xml files
                For Each orderFile As String In My.Computer.FileSystem.GetFiles(vwConnection.LocalInDir, FileIO.SearchOption.SearchTopLevelOnly, "*.xml")
                    Try
                        errors.Clear()
                        ERR_LNO = 0
                        specialInstructions = String.Empty
                        invalidCustomer = False

                        With dst
                            .Tables("DETJOBM1").Clear()
                            .Tables("DETJOBM2").Clear()
                            .Tables("DETJOBM3").Clear()
                            .Tables("DETJOBM4").Clear()
                            .Tables("DETJOBIE").Clear()

                            .Tables("SOTORDR1").Clear()
                            .Tables("SOTORDR5").Clear()
                        End With

                        Try
                            For Each tbl As DataTable In vwXmlDataset.Tables
                                tbl.Rows.Clear()
                                tbl.BeginLoadData()
                            Next
                            vwXmlDataset.ReadXml(orderFile)
                            For Each tbl As DataTable In vwXmlDataset.Tables
                                tbl.EndLoadData()
                            Next
                        Catch ex As Exception
                            My.Computer.FileSystem.MoveFile(orderFile, errorsDirectory & My.Computer.FileSystem.GetName(orderFile), True)
                            Dim note As String = "DEL Vision Web file (" & orderFile & ") caused the following error: " & ex.Message & Environment.NewLine
                            note &= " File placed in directory: " & errorsDirectory & "." & Environment.NewLine
                            note &= " The job was not saved!!" & Environment.NewLine
                            emailErrors(String.Empty, 1, note, "DEL_VWEB")
                            Continue For
                        End Try

                        RX_SPECTACLE_Id = String.Empty
                        HEADER_Id = String.Empty
                        CUSTOMER_Id = String.Empty
                        ACCOUNTS_Id = String.Empty
                        ACCOUNT_Id = String.Empty
                        SP_EQUIPMENT_Id = String.Empty
                        PATIENT_ID = String.Empty
                        POSITION_ID = String.Empty
                        PRESCRIPTION_ID = String.Empty
                        LENS_ID = String.Empty
                        TREATMENTS_ID = String.Empty

                        For Each rowRX_SPECTACLE As DataRow In vwXmlDataset.Tables("RX_SPECTACLE").Rows

                            Dim rowDETJOBM1 As DataRow = dst.Tables("DETJOBM1").NewRow
                            RX_SPECTACLE_Id = rowRX_SPECTACLE.Item("RX_SPECTACLE_Id") & String.Empty
                            HEADER_Id = String.Empty
                            CUSTOMER_Id = String.Empty
                            ACCOUNTS_Id = String.Empty
                            ACCOUNT_Id = String.Empty
                            SP_EQUIPMENT_Id = String.Empty
                            PATIENT_ID = String.Empty
                            POSITION_ID = String.Empty
                            PRESCRIPTION_ID = String.Empty
                            LENS_ID = String.Empty
                            TREATMENTS_ID = String.Empty

                            JOB_NO = ABSolution.ASCMAIN1.Next_Control_No("DETJOBM1.JOB_NO", 1)
                            ORDR_NO = ABSolution.ASCMAIN1.Next_Control_No("SOTORDR1.ORDR_NO", 1)
                            rowDETJOBM1.Item("JOB_NO") = JOB_NO
                            rowDETJOBM1.Item("ORDR_NO") = ORDR_NO
                            dst.Tables("DETJOBM1").Rows.Add(rowDETJOBM1)

                            ' Default values taken from DETJOBM1
                            rowDETJOBM1.Item("JOB_STATUS") = "O"
                            rowDETJOBM1.Item("ORDR_SOURCE") = ORDR_SOURCE
                            rowDETJOBM1.Item("JOB_TYPE_CODE") = "O" ' Original
                            rowDETJOBM1.Item("ORDR_DATE") = DateTime.Now.ToString("MM/dd/yyyy")
                            rowDETJOBM1.Item("ORDR_CUST_PO") = "DEL/VW/"
                            rowDETJOBM1.Item("COMMENT_LAB") = String.Empty
                            rowDETJOBM1.Item("JOB_STATUS") = "O"

                            rowDETJOBM1.Item("INIT_DATE") = DateTime.Now
                            rowDETJOBM1.Item("INIT_OPER") = ABSolution.ASCMAIN1.USER_ID
                            rowDETJOBM1.Item("LAST_DATE") = DateTime.Now
                            rowDETJOBM1.Item("LAST_OPER") = ABSolution.ASCMAIN1.USER_ID

                            rowDETJOBM1.Item("LENS_ORDER") = "B"
                            rowDETJOBM1.Item("FINISHED") = "U"
                            rowDETJOBM1.Item("FRAME_STATUS") = "N"
                            rowDETJOBM1.Item("TINT_CODE") = "NONE"
                            rowDETJOBM1.Item("FRAME_TYPE_CODE") = "NONE"
                            rowDETJOBM1.Item("TRACE_FROM") = "N"
                            rowDETJOBM1.Item("POLISHING") = "0"
                            rowDETJOBM1.Item("ORDR_CALLER_NAME") = "Service"
                            rowDETJOBM1.Item("USE_THINNING_PRISM") = "1"
                            rowDETJOBM1.Item("COLOR_TYPE") = "C"
                            rowDETJOBM1.Item("RX_PRISM") = "0"
                            rowDETJOBM1.Item("EDGING") = "0"
                            rowDETJOBM1.Item("NO_FREIGHT") = "0"
                            rowDETJOBM1.Item("INV_FREIGHT") = 0
                            rowDETJOBM1.Item("LIST_PRICE") = 0
                            rowDETJOBM1.Item("INV_TOTAL_AMOUNT") = 0
                            rowDETJOBM1.Item("OPS_YYYYPP") = "" ' ABSolution.ASCMAIN1.CYP
                            rowDETJOBM1.Item("USE_DISC_PCT") = "0"
                            rowDETJOBM1.Item("MIRROR_COATING") = "0"
                            rowDETJOBM1.Item("AR_BACKSIDE_ONLY") = "0"
                            rowDETJOBM1.Item("INV_SALES") = 0
                            rowDETJOBM1.Item("JOB_HOLD_LAB") = "0"
                            rowDETJOBM1.Item("JOB_HOLD_INV") = "0"
                            rowDETJOBM1.Item("JOB_IN_QUEUE") = "0"
                            rowDETJOBM1.Item("JOB_REQUIRES_REVIEW") = "0"
                            rowDETJOBM1.Item("WRAP_EDGE") = "0"
                            rowDETJOBM1.Item("CUSTOM_FRAME_NEW") = "0"
                            rowDETJOBM1.Item("WRAP_EDGE_SPORT") = "0"
                            rowDETJOBM1.Item("FOG_FREE") = "0"
                            rowDETJOBM1.Item("UNCUT_TINTABLE") = "0"
                            rowDETJOBM1.Item("JOB_NO_BACKSIDE_COAT") = "0"
                            rowDETJOBM1.Item("LDS_IND") = "0"
                            rowDETJOBM1.Item("LMS_IND") = "0"
                            rowDETJOBM1.Item("TKT_IND") = "0"
                            rowDETJOBM1.Item("BLANK_SELECTION") = "N"
                            rowDETJOBM1.Item("BALANCE_LENS") = "0"
                            rowDETJOBM1.Item("TKT_PRINT_COUNT") = 0
                            ' As per Maria place in Queue
                            rowDETJOBM1.Item("JOB_IN_QUEUE") = "1"

                            creationDate = (rowRX_SPECTACLE.Item("Creation_date") & String.Empty).ToString.Replace("T", Space(1))
                            If IsDate(creationDate) Then
                                rowDETJOBM1.Item("ORDR_DATE") = CDate(creationDate).ToString("dd-MMM-yyyy")
                            End If

                            For Each rowHeader As DataRow In vwXmlDataset.Tables("HEADER").Select("RX_SPECTACLE_Id = " & RX_SPECTACLE_Id)
                                HEADER_Id = rowHeader.Item("HEADER_Id") & String.Empty
                                rowDETJOBM1.Item("ORDR_CUST_PO") = TruncateField("DEL/VW/" & rowRX_SPECTACLE.Item("Id"), "DETJOBM1", "ORDR_CUST_PO")
                                rowDETJOBM1.Item("ORDR_CALLER_NAME") = TruncateField(rowHeader.Item("OrderPlacedBy") & String.Empty, "DETJOBM1", "ORDR_CALLER_NAME")
                                rowDETJOBM1.Item("ORDR_SOURCE") = ORDR_SOURCE

                                CUSTOMER_Id = String.Empty
                                For Each rowCustomer As DataRow In vwXmlDataset.Tables("CUSTOMER").Select("HEADER_Id = " & HEADER_Id, "CUSTOMER_ID")
                                    CUSTOMER_Id = rowCustomer.Item("CUSTOMER_Id") & String.Empty

                                    For Each rowACCOUNTS As DataRow In vwXmlDataset.Tables("ACCOUNTS").Select("CUSTOMER_Id = " & CUSTOMER_Id, "ACCOUNTS_ID")
                                        ACCOUNTS_Id = rowACCOUNTS.Item("ACCOUNTS_ID") & String.Empty
                                        CUST_CODE = String.Empty
                                        CUST_SHIP_TO_NO = String.Empty

                                        For Each rowACCOUNT As DataRow In vwXmlDataset.Tables("ACCOUNT").Select("ACCOUNTS_ID = " & ACCOUNTS_Id, "ACCOUNT_ID")

                                            Select Case (rowACCOUNT.Item("Class") & String.Empty).ToString.Trim.ToUpper
                                                Case "BIL"
                                                    ACCOUNT_Id = rowACCOUNT.Item("ACCOUNT_Id") & String.Empty
                                                    If ACCOUNT_Id.Length > 0 AndAlso vwXmlDataset.Tables("ADDRESS").Select("ACCOUNT_Id = " & ACCOUNT_Id).Length > 0 Then
                                                        rowData = vwXmlDataset.Tables("ADDRESS").Select("ACCOUNT_Id = " & ACCOUNT_Id)(0)

                                                        rowSOTORDR5 = dst.Tables("SOTORDR5").NewRow
                                                        rowSOTORDR5.Item("ORDR_NO") = ORDR_NO
                                                        rowSOTORDR5.Item("CUST_ADDR_TYPE") = "BT"
                                                        dst.Tables("SOTORDR5").Rows.Add(rowSOTORDR5)

                                                        rowSOTORDR5.Item("CUST_NAME") = String.Empty
                                                        rowSOTORDR5.Item("CUST_ADDR1") = TruncateField((rowData.Item("Street_Number") & String.Empty & " " & rowData.Item("Street_Name") & String.Empty).ToString.Trim, "SOTORDR5", "CUST_ADDR1")
                                                        rowSOTORDR5.Item("CUST_ADDR2") = TruncateField(rowData.Item("Suite") & String.Empty, "SOTORDR5", "CUST_ADDR2")
                                                        rowSOTORDR5.Item("CUST_CITY") = TruncateField(rowData.Item("City") & String.Empty, "SOTORDR5", "CUST_CITY")
                                                        rowSOTORDR5.Item("CUST_STATE") = TruncateField(rowData.Item("State") & String.Empty, "SOTORDR5", "CUST_STATE")
                                                        rowSOTORDR5.Item("CUST_ZIP_CODE") = TruncateField(rowData.Item("ZipCode") & String.Empty, "SOTORDR5", "CUST_ZIP_CODE")
                                                        Telephone = rowData.Item("TEL") & String.Empty
                                                        Telephone = FormatTelePhone(Telephone)
                                                        rowSOTORDR5.Item("CUST_PHONE") = TruncateField(Telephone, "SOTORDR5", "CUST_PHONE")
                                                        rowSOTORDR5.Item("CUST_COUNTRY") = TruncateField(rowData.Item("Country") & String.Empty, "SOTORDR5", "CUST_COUNTRY")
                                                        rowSOTORDR5.Item("CUST_FAX") = String.Empty
                                                        rowSOTORDR5.Item("CUST_EMAIL") = String.Empty
                                                    End If

                                                Case "SHP"
                                                    ACCOUNT_Id = rowACCOUNT.Item("ACCOUNT_Id") & String.Empty
                                                    If ACCOUNT_Id.Length > 0 AndAlso vwXmlDataset.Tables("ADDRESS").Select("ACCOUNT_Id = " & ACCOUNT_Id).Length > 0 Then
                                                        rowData = vwXmlDataset.Tables("ADDRESS").Select("ACCOUNT_Id = " & ACCOUNT_Id)(0)

                                                        rowSOTORDR5 = dst.Tables("SOTORDR5").NewRow
                                                        rowSOTORDR5.Item("ORDR_NO") = ORDR_NO
                                                        rowSOTORDR5.Item("CUST_ADDR_TYPE") = "ST"
                                                        dst.Tables("SOTORDR5").Rows.Add(rowSOTORDR5)

                                                        rowSOTORDR5.Item("CUST_NAME") = String.Empty
                                                        rowSOTORDR5.Item("CUST_ADDR1") = TruncateField((rowData.Item("Street_Number") & String.Empty & " " & rowData.Item("Street_Name") & String.Empty).ToString.Trim, "SOTORDR5", "CUST_ADDR1")
                                                        rowSOTORDR5.Item("CUST_ADDR2") = TruncateField(rowData.Item("Suite") & String.Empty, "SOTORDR5", "CUST_ADDR2")
                                                        rowSOTORDR5.Item("CUST_CITY") = TruncateField(rowData.Item("City") & String.Empty, "SOTORDR5", "CUST_CITY")
                                                        rowSOTORDR5.Item("CUST_STATE") = TruncateField(rowData.Item("State") & String.Empty, "SOTORDR5", "CUST_STATE")
                                                        rowSOTORDR5.Item("CUST_ZIP_CODE") = TruncateField(rowData.Item("Zipcode") & String.Empty, "SOTORDR5", "CUST_ZIP_CODE")
                                                        Telephone = (rowData.Item("TEL") & String.Empty).trim
                                                        Telephone = FormatTelePhone(Telephone)
                                                        rowSOTORDR5.Item("CUST_PHONE") = TruncateField(Telephone, "SOTORDR5", "CUST_PHONE")
                                                        rowSOTORDR5.Item("CUST_FAX") = String.Empty
                                                        rowSOTORDR5.Item("CUST_EMAIL") = String.Empty
                                                        rowSOTORDR5.Item("CUST_COUNTRY") = TruncateField(rowData.Item("Country") & String.Empty, "SOTORDR5", "CUST_COUNTRY")

                                                        CUST_CODE = rowACCOUNT.Item("Name") & String.Empty
                                                        If ABSolution.ASCMAIN1.DBS_COMPANY.Trim.ToUpper = "TST" AndAlso CUST_CODE = "1234567" Then
                                                            CUST_CODE = "77780"
                                                        End If
                                                        If CUST_CODE.Contains("-") Then
                                                            CUST_SHIP_TO_NO = Split(CUST_CODE, "-")(1)
                                                            CUST_CODE = Split(CUST_CODE, "-")(0)
                                                        End If

                                                        CUST_CODE = ABSolution.ASCMAIN1.Format_Field(CUST_CODE, "CUST_CODE")
                                                        If CUST_CODE.Length = 0 Then
                                                            CUST_CODE = rowACCOUNT.Item("Name") & String.Empty
                                                        End If
                                                        CUST_CODE = TruncateField(CUST_CODE, "SOTORDR5", "CUST_CODE")
                                                        If CUST_SHIP_TO_NO.Length > 0 Then
                                                            CUST_SHIP_TO_NO = ABSolution.ASCMAIN1.Format_Field(CUST_SHIP_TO_NO, "CUST_SHIP_TO_NO")
                                                            CUST_SHIP_TO_NO = TruncateField(CUST_SHIP_TO_NO, "DETJOBM1", "CUST_SHIP_TO_NO")
                                                        End If
                                                    End If
                                            End Select
                                        Next
                                    Next

                                    For Each rowSP_EQUIPMENT As DataRow In vwXmlDataset.Tables("SP_EQUIPMENT").Select("RX_SPECTACLE_Id = " & RX_SPECTACLE_Id)
                                        SP_EQUIPMENT_Id = rowSP_EQUIPMENT.Item("SP_EQUIPMENT_Id") & String.Empty

                                        If vwXmlDataset.Tables("PATIENT").Select("SP_EQUIPMENT_Id = " & SP_EQUIPMENT_Id).Length > 0 Then
                                            rowData = vwXmlDataset.Tables("PATIENT").Select("SP_EQUIPMENT_Id = " & SP_EQUIPMENT_Id)(0)
                                            tempString = (rowData.Item("LASTNAME") & ", " & rowData.Item("FIRSTNAME") & String.Empty).ToString.Trim
                                            tempString = tempString.Trim.ToUpper
                                            rowDETJOBM1.Item("PATIENT_NAME") = TruncateField(tempString, "DETJOBM1", "PATIENT_NAME")
                                            PATIENT_ID = (rowData.Item("PATIENT_ID") & String.Empty).ToString.Trim
                                        End If

                                        If vwXmlDataset.Tables("FRAME").Select("SP_EQUIPMENT_Id = " & SP_EQUIPMENT_Id).Length > 0 Then
                                            Dim rowFrame As DataRow = vwXmlDataset.Tables("FRAME").Select("SP_EQUIPMENT_Id = " & SP_EQUIPMENT_Id)(0)
                                            Dim FRAME_TYPE_CODE As String = rowFrame.Item("CONVERT") & String.Empty
                                            Dim ORDR_MESSAGE_LAB As String = rowFrame.Item("COMMENT") & String.Empty
                                            rowDETJOBM1.Item("ORDR_MESSAGE_LAB") = TruncateField(ORDR_MESSAGE_LAB, "DETJOBM1", "ORDR_MESSAGE_LAB")

                                            Dim rowDETFRAM1 As DataRow = ABSolution.ASCDATA1.GetDataRow("Select * from DETFRAM1 where upper(FRAME_TYPE_CODE) = :PARM1", "V", New Object() {FRAME_TYPE_CODE.ToUpper})
                                            If rowDETFRAM1 IsNot Nothing Then
                                                rowDETJOBM1.Item("FRAME_TYPE_CODE") = rowDETFRAM1.Item("FRAME_TYPE_CODE") & String.Empty
                                            End If

                                            'PFL: Package, 
                                            'TBP: to be purchase, 
                                            'TC to come, 
                                            'PRES : Presize, 
                                            'NONE, 
                                            'USF : Uncut - Supply Frame, 
                                            'ESF : Edged - Supply Frame
                                            Dim FRAME_STATUS As String = (rowFrame.Item("ACTION") & String.Empty).ToString.Trim.ToUpper
                                            Select Case FRAME_STATUS
                                                Case "TC"
                                                    rowDETJOBM1.Item("FRAME_STATUS") = "C"
                                                    rowDETJOBM1.Item("FINISHED") = "F"
                                                    rowDETJOBM1.Item("EDGING") = "1"
                                                Case "ESF"
                                                    rowDETJOBM1.Item("FRAME_STATUS") = "S"
                                                    rowDETJOBM1.Item("FINISHED") = "F"
                                                    rowDETJOBM1.Item("EDGING") = "1"
                                                Case "PRES"
                                                    rowDETJOBM1.Item("FRAME_STATUS") = "S"
                                                    rowDETJOBM1.Item("FINISHED") = "F"
                                                    rowDETJOBM1.Item("EDGING") = "1"
                                                Case Else
                                                    rowDETJOBM1.Item("FRAME_STATUS") = "N"
                                            End Select

                                            ' Default from here, then overwrite with data from SHAPE
                                            rowDETJOBM1.Item("FRAME_A_WIDTH") = Val(rowFrame.Item("A") & String.Empty)
                                            rowDETJOBM1.Item("FRAME_B_HEIGHT") = Val(rowFrame.Item("B") & String.Empty)
                                            rowDETJOBM1.Item("FRAME_DBL_BRIDGE") = Val(rowFrame.Item("DBL") & String.Empty)
                                            rowDETJOBM1.Item("FRAME_ED_DIAGONAL") = Val(rowFrame.Item("ED") & String.Empty)

                                            rowDETJOBM1.Item("FRAME_MFG") = rowFrame.Item("BRAND") & String.Empty
                                            rowDETJOBM1.Item("FRAME_MODEL_NO") = rowFrame.Item("MODEL") & String.Empty
                                            rowDETJOBM1.Item("FRAME_SIZE") = rowFrame.Item("EYESIZE") & String.Empty
                                            rowDETJOBM1.Item("FRAME_COLOR") = rowFrame.Item("COLOR") & String.Empty

                                        End If

                                        ' Create the detjobm3 records
                                        Dim rowDETJOBM3 As DataRow = Nothing

                                        If dst.Tables("DETJOBM3").Select("JOB_NO = '" & JOB_NO & "' AND RL = 'R'").Length = 0 Then
                                            rowDETJOBM3 = dst.Tables("DETJOBM3").NewRow()
                                            rowDETJOBM3.Item("JOB_NO") = JOB_NO
                                            rowDETJOBM3.Item("RL") = "R"
                                            dst.Tables("DETJOBM3").Rows.Add(rowDETJOBM3)
                                        End If

                                        If dst.Tables("DETJOBM3").Select("JOB_NO = '" & JOB_NO & "' AND RL = 'L'").Length = 0 Then
                                            rowDETJOBM3 = dst.Tables("DETJOBM3").NewRow()
                                            rowDETJOBM3.Item("JOB_NO") = JOB_NO
                                            rowDETJOBM3.Item("RL") = "L"
                                            dst.Tables("DETJOBM3").Rows.Add(rowDETJOBM3)
                                        End If

                                        For Each rowPOSITION As DataRow In vwXmlDataset.Tables("POSITION").Select("SP_EQUIPMENT_Id = " & SP_EQUIPMENT_Id)
                                            POSITION_ID = rowPOSITION.Item("POSITION_ID") & String.Empty
                                            PRESCRIPTION_ID = String.Empty
                                            LENS_ID = String.Empty
                                            TREATMENTS_ID = String.Empty

                                            If Not "RL".Contains(rowPOSITION.Item("EYE") & String.Empty) Then
                                                rowDETJOBM1.Item("LENS_ORDER") = "X"
                                                Continue For
                                            End If

                                            rowDETJOBM3 = dst.Tables("DETJOBM3").Select("JOB_NO = '" & JOB_NO & "' AND RL = '" & rowPOSITION.Item("EYE") & "'")(0)
                                            'rowDETJOBM3.Item("JOB_NO") = JOB_NO
                                            'rowDETJOBM3.Item("RL") = rowPOSITION.Item("EYE") & String.Empty
                                            rowDETJOBM3.Item("MONO_PD") = Val(rowPOSITION.Item("FAR_HALF_PD") & String.Empty)
                                            rowDETJOBM3.Item("FITTING_HEIGHT") = Val(rowPOSITION.Item("BOXING_HEIGHT") & String.Empty)
                                            'dst.Tables("DETJOBM3").Rows.Add(rowDETJOBM3)

                                            ' Need to see if Both or a single lens
                                            If rowDETJOBM1.Item("LENS_ORDER") = "B" Then
                                                If rowDETJOBM3.Item("RL") & String.Empty = "L" Or rowDETJOBM3.Item("RL") & String.Empty = "R" Then
                                                    rowDETJOBM1.Item("LENS_ORDER") = rowDETJOBM3.Item("RL") & String.Empty
                                                End If
                                            ElseIf rowDETJOBM3.Item("RL") & String.Empty = "L" Or rowDETJOBM3.Item("RL") & String.Empty = "R" Then
                                                If rowDETJOBM1.Item("LENS_ORDER") <> "B" AndAlso _
                                                        rowDETJOBM1.Item("LENS_ORDER") <> rowDETJOBM3.Item("RL") & String.Empty Then
                                                    rowDETJOBM1.Item("LENS_ORDER") = "B"
                                                End If
                                            End If

                                            If vwXmlDataset.Tables("SHAPE").Select("POSITION_ID = " & POSITION_ID).Length > 0 Then
                                                Dim rowShape As DataRow = vwXmlDataset.Tables("SHAPE").Select("POSITION_ID = " & POSITION_ID)(0)

                                                rowDETJOBM1.Item("FRAME_A_WIDTH") = Val(rowShape.Item("A") & String.Empty)
                                                rowDETJOBM1.Item("FRAME_B_HEIGHT") = Val(rowShape.Item("B") & String.Empty)
                                                rowDETJOBM1.Item("FRAME_DBL_BRIDGE") = Val(rowShape.Item("HALF_DBL") & String.Empty) * 2
                                                rowDETJOBM1.Item("FRAME_ED_DIAGONAL") = Val(rowShape.Item("E") & String.Empty)
                                            End If

                                            ' As per maria/vince 9/5/2012
                                            If Val(rowDETJOBM1.Item("FRAME_ED_DIAGONAL") & String.Empty) < Val(rowDETJOBM1.Item("FRAME_A_WIDTH") & String.Empty) Then
                                                rowDETJOBM1.Item("FRAME_ED_DIAGONAL") = Val(rowDETJOBM1.Item("FRAME_A_WIDTH") & String.Empty)
                                            End If

                                            If vwXmlDataset.Tables("PRESCRIPTION").Select("POSITION_ID = " & POSITION_ID).Length > 0 Then
                                                rowData = vwXmlDataset.Tables("PRESCRIPTION").Select("POSITION_ID = " & POSITION_ID)(0)
                                                PRESCRIPTION_ID = (rowData.Item("PRESCRIPTION_ID") & String.Empty).ToString.Trim
                                                rowDETJOBM3.Item("SPHERE") = Val(rowData.Item("SPHERE") & String.Empty)
                                            End If

                                            If vwXmlDataset.Tables("CYLINDER").Select("PRESCRIPTION_ID = " & PRESCRIPTION_ID).Length > 0 Then
                                                rowData = vwXmlDataset.Tables("CYLINDER").Select("PRESCRIPTION_ID = " & PRESCRIPTION_ID)(0)
                                                rowDETJOBM3.Item("CYLINDER") = Val(rowData.Item("VALUE") & String.Empty)
                                                rowDETJOBM3.Item("AXIS") = Val(rowData.Item("AXIS") & String.Empty)
                                            End If

                                            If vwXmlDataset.Tables("ADDITION").Select("PRESCRIPTION_ID = " & PRESCRIPTION_ID).Length > 0 Then
                                                rowData = vwXmlDataset.Tables("ADDITION").Select("PRESCRIPTION_ID = " & PRESCRIPTION_ID)(0)
                                                rowDETJOBM3.Item("ADD_POWER") = Val(rowData.Item("VALUE") & String.Empty)
                                            End If

                                            For Each rowPRISM As DataRow In vwXmlDataset.Tables("PRISM").Select("PRESCRIPTION_id = " & PRESCRIPTION_ID)
                                                Select Case (rowPRISM.Item("AXISSTR") & String.Empty).ToString.Trim.ToUpper
                                                    Case "IN", "OUT"
                                                        rowDETJOBM3.Item("PRISM_IN_AXIS") = (rowPRISM.Item("AXISSTR") & String.Empty).ToString.Trim.ToUpper.Substring(0, 1)
                                                        rowDETJOBM3.Item("PRISM_IN") = Val(rowPRISM.Item("VALUE") & String.Empty)
                                                    Case "UP", "DOWN"
                                                        rowDETJOBM3.Item("PRISM_UP_AXIS") = (rowPRISM.Item("AXISSTR") & String.Empty).ToString.Trim.ToUpper.Substring(0, 1)
                                                        rowDETJOBM3.Item("PRISM_UP") = Val(rowPRISM.Item("VALUE") & String.Empty)
                                                    Case Else
                                                        specialInstructions &= "Error on Prism data" & Environment.NewLine
                                                        Continue For
                                                End Select
                                                rowDETJOBM1.Item("RX_PRISM") = "1"
                                            Next

                                            If vwXmlDataset.Tables("CALCULATIONS").Select("POSITION_ID = " & POSITION_ID).Length > 0 Then
                                                rowData = vwXmlDataset.Tables("CALCULATIONS").Select("POSITION_ID = " & POSITION_ID)(0)
                                                For ictr As Int16 = 0 To rowData.Table.Columns.Count - 1
                                                    Select Case rowData.Table.Columns(ictr).ColumnName.ToUpper
                                                        Case "POSITION_ID", "CALCULATIONS_ID"
                                                            ' Nothing
                                                        Case Else
                                                            If rowData.Item(rowData.Table.Columns(ictr).ColumnName) & String.Empty <> String.Empty Then
                                                                specialInstructions &= rowPOSITION.Item("EYE") & String.Empty & " eye:"
                                                                specialInstructions &= rowData.Table.Columns(ictr).ColumnName & "="
                                                                specialInstructions &= rowData.Item(rowData.Table.Columns(ictr).ColumnName)
                                                                specialInstructions &= Environment.NewLine
                                                            End If
                                                    End Select
                                                Next
                                            End If

                                            If vwXmlDataset.Tables("LENS").Select("POSITION_ID = " & POSITION_ID).Length > 0 Then
                                                rowData = vwXmlDataset.Tables("LENS").Select("POSITION_ID = " & POSITION_ID)(0)
                                                LENS_ID = (rowData.Item("LENS_ID") & String.Empty).ToString.Trim

                                                If vwXmlDataset.Tables("DESIGN").Select("LENS_ID = " & LENS_ID).Length > 0 Then
                                                    rowData = vwXmlDataset.Tables("DESIGN").Select("LENS_ID = " & LENS_ID)(0)
                                                    rowDETJOBM1.Item("LENS_DESIGN_CODE") = (rowData.Item("CONVERT") & String.Empty).ToString.Trim
                                                End If

                                                If vwXmlDataset.Tables("THICKNESS").Select("LENS_ID = " & LENS_ID).Length > 0 Then
                                                    rowData = vwXmlDataset.Tables("THICKNESS").Select("LENS_ID = " & LENS_ID)(0)
                                                    specialInstructions &= rowPOSITION.Item("EYE") & String.Empty & " eye thickness:"
                                                    specialInstructions &= (rowData.Item("TYPE") & "=" & rowData.Item("VALUE")).ToString.Trim
                                                    specialInstructions &= Environment.NewLine
                                                End If

                                                If vwXmlDataset.Tables("MATERIAL").Select("LENS_ID = " & LENS_ID).Length > 0 Then
                                                    rowData = vwXmlDataset.Tables("MATERIAL").Select("LENS_ID = " & LENS_ID)(0)
                                                    tempString = (rowData.Item("CONVERT") & String.Empty).ToString.Trim
                                                    If tempString.Split(":").Length >= 2 Then
                                                        rowDETJOBM1.Item("MATL_CODE") = tempString.Split(":")(0)
                                                        rowDETJOBM1.Item("COLOR_CODE") = tempString.Split(":")(1)
                                                    Else
                                                        rowDETJOBM1.Item("MATL_CODE") = tempString
                                                    End If
                                                End If

                                                If vwXmlDataset.Tables("TREATMENTS").Select("LENS_ID = " & LENS_ID).Length > 0 Then
                                                    rowData = vwXmlDataset.Tables("TREATMENTS").Select("LENS_ID = " & LENS_ID)(0)
                                                    TREATMENTS_ID = (rowData.Item("TREATMENTS_ID") & String.Empty).ToString.Trim

                                                    For Each rowTREATMENT As DataRow In vwXmlDataset.Tables("TREATMENT").Select("TREATMENTS_ID = " & TREATMENTS_ID)
                                                        tempString = rowTREATMENT.Item("CONVERT") & String.Empty

                                                        Dim rowDETJOBVW As DataRow = ABSolution.ASCDATA1.GetDataRow("SELECT * FROM DETJOBVW WHERE ACTION_CODE = :PARM1 AND ACTION_TYPE = :PARM2", _
                                                                                                                    "VV", New Object() {tempString, "TREATMENT"})
                                                        If rowDETJOBVW IsNot Nothing Then
                                                            For Each field As String In New String() {"AR_COATING", "AR_BACKSIDE_ONLY", "COLOR_CODE", "COLOR_TYPE", _
                                                                                                      "FRAME_STATUS", "FRAME_TYPE_CODE", "LENS_DESIGN_CODE", _
                                                                                                      "LENS_DESIGNER_CODE", "MATL_CODE", "MIRROR_COATING", _
                                                                                                      "POLISHING", "TINT_CODE", "TINT_COLOR", "WRAP_EDGE", _
                                                                                                      "MIRROR_COATING_COLOR", "FINISHED", "FOG_FREE"}
                                                                If rowDETJOBVW.Item(field) & String.Empty <> String.Empty Then
                                                                    rowDETJOBM1.Item(field) = rowDETJOBVW.Item(field) & String.Empty
                                                                End If
                                                            Next ' For Each field 
                                                        End If
                                                    Next ' For Each rowTREATMENT
                                                End If
                                            End If
                                        Next ' For Each rowPOSITION

                                        ' Trace data
                                        If vwXmlDataset.Tables("OTHER").Select("SP_EQUIPMENT_Id = " & SP_EQUIPMENT_Id).Length > 0 Then
                                            Dim OTHER_ID As String = vwXmlDataset.Tables("OTHER").Select("SP_EQUIPMENT_Id = " & SP_EQUIPMENT_Id)(0).Item("OTHER_ID") & String.Empty
                                            If vwXmlDataset.Tables("OMA").Select("OTHER_ID = " & OTHER_ID).Length > 0 Then
                                                Dim traceData As String = vwXmlDataset.Tables("OMA").Select("OTHER_ID = " & OTHER_ID)(0).Item("DATA") & String.Empty
                                                Try
                                                    'alter trace data to appropriate format before writing
                                                    Dim traceLines() As String = traceData.Split(" ")

                                                    'write trace file
                                                    Using traceFileToSave As New System.IO.FileStream("\\192.168.130.201\Shared\DEL\TRC\" & JOB_NO & ".trc", IO.FileMode.Create)
                                                        Using traceFileStream As New System.IO.StreamWriter(traceFileToSave)
                                                            For Each line As String In traceLines
                                                                traceFileStream.WriteLine(line)
                                                            Next
                                                            traceFileStream.Close()
                                                            traceFileToSave.Close()
                                                        End Using
                                                    End Using

                                                    rowDETJOBM1.Item("TRACE_FROM") = "T"

                                                Catch ex As Exception
                                                    errors.Add(New DelJobService.DelJobValidationError("TraceData", "Error: " & ex.Message))
                                                End Try
                                            End If
                                        End If
                                    Next ' For Each rowSP_EQUIPMENT
                                Next 'For Each rowACCOUNT
                            Next ' For Each customer

                            ' Convert the color code
                            Dim COLOR_CODE As String = rowDETJOBM1.Item("COLOR_CODE") & String.Empty
                            Dim MATL_CODE As String = rowDETJOBM1.Item("MATL_CODE") & String.Empty
                            Dim LENS_DESIGN_CODE As String = rowDETJOBM1.Item("LENS_DESIGN_CODE") & String.Empty
                            Dim rowCOLORCODE As DataRow = ABSolution.ASCDATA1.GetDataRow(sqlColorCode, "VVV", New Object() {LENS_DESIGN_CODE, MATL_CODE, COLOR_CODE})
                            If rowCOLORCODE IsNot Nothing Then
                                rowDETJOBM1.Item("COLOR_CODE") = rowCOLORCODE.Item("COLOR_CODE") & String.Empty
                                If rowCOLORCODE.Item("COLOR_CODE") = "CLEAR" Then
                                    rowDETJOBM1.Item("COLOR_TYPE") = "C"
                                ElseIf rowCOLORCODE.Item("POLARIZED") & String.Empty = "1" Then '
                                    rowDETJOBM1.Item("COLOR_TYPE") = "P"
                                Else
                                    rowDETJOBM1.Item("COLOR_TYPE") = "T"
                                End If
                            End If

                            ' Corridor Length
                            Dim FITTING_HEIGHT As Double
                            FITTING_HEIGHT = Val(dst.Tables("DETJOBM3").Compute("MAX(FITTING_HEIGHT)", String.Empty) & String.Empty)

                            Dim SQL = "Select CORRIDOR_LENGTH from DETDSGN2 " _
                                & " where LENS_DESIGN_CODE = :PARM1" _
                                & " and MIN_FITTING_HEIGHT = " _
                                & " (Select Max (MIN_FITTING_HEIGHT) from DETDSGN2 " _
                                & " where LENS_DESIGN_CODE = :PARM2" _
                                & " and MIN_FITTING_HEIGHT <= :PARM3)"

                            Dim CORRIDOR_LENGTH As String = ABSolution.ASCDATA1.GetDataValue(SQL, "VVN", New Object() {LENS_DESIGN_CODE, LENS_DESIGN_CODE, FITTING_HEIGHT})
                            If CORRIDOR_LENGTH <> String.Empty Then
                                rowDETJOBM1.Item("CORRIDOR_LENGTH") = CORRIDOR_LENGTH
                            End If

                            rowDETJOBM1.Item("COMMENT_LAB") = specialInstructions

                            rowARTCUST1 = baseClass.LookUp("ARTCUST1", CUST_CODE)
                            rowARTCUST2 = baseClass.LookUp("ARTCUST2", New String() {CUST_CODE, CUST_SHIP_TO_NO})
                            rowARTCUST3 = baseClass.LookUp("ARTCUST3", CUST_CODE)

                            ' If not a valid customer then place file in Errors File
                            If rowARTCUST1 Is Nothing OrElse (CUST_SHIP_TO_NO.Length > 0 AndAlso rowARTCUST2 Is Nothing) Then
                                My.Computer.FileSystem.MoveFile(orderFile, errorsDirectory & My.Computer.FileSystem.GetName(orderFile), True)
                                Dim note As String = "Unknown Customer Code: " & CUST_CODE
                                If CUST_SHIP_TO_NO.Length > 0 AndAlso rowARTCUST2 Is Nothing Then
                                    note = ", Ship to: " & CUST_SHIP_TO_NO
                                End If
                                note &= " in file " & My.Computer.FileSystem.GetName(orderFile) & Environment.NewLine
                                note &= " File placed in directory: " & errorsDirectory & Environment.NewLine
                                note &= " The job was not saved!!" & Environment.NewLine
                                emailErrors(String.Empty, errors.Count, note, "DEL_VWEB")
                                invalidCustomer = True
                                Continue For
                            End If

                            If rowARTCUST3 IsNot Nothing Then
                                rowDETJOBM1.Item("SHIP_VIA_CODE") = rowARTCUST3.Item("SHIP_VIA_CODE")

                                If CUST_SHIP_TO_NO <> String.Empty Then
                                    If rowARTCUST2.Item("CUST_SHIP_TO_SHIP_VIA_CODE") & "" <> "" Then
                                        rowDETJOBM1.Item("SHIP_VIA_CODE") = rowARTCUST2.Item("CUST_SHIP_TO_SHIP_VIA_CODE")
                                    End If
                                End If
                            End If

                            rowDETJOBM1.Item("CUST_CODE") = CUST_CODE
                            If rowARTCUST1 IsNot Nothing Then
                                rowDETJOBM1.Item("CUST_NAME") = rowARTCUST1.Item("CUST_NAME") & String.Empty

                                ' Update Sotordr5 BT Record
                                If dst.Tables("SOTORDR5").Select("ORDR_NO = '" & ORDR_NO & "' and CUST_ADDR_TYPE = 'BT'").Length > 0 Then
                                    rowSOTORDR5 = dst.Tables("SOTORDR5").Select("ORDR_NO = '" & ORDR_NO & "' and CUST_ADDR_TYPE = 'BT'")(0)
                                Else
                                    rowSOTORDR5 = dst.Tables("SOTORDR5").NewRow
                                    rowSOTORDR5.Item("ORDR_NO") = ORDR_NO
                                    rowSOTORDR5.Item("CUST_ADDR_TYPE") = "BT"
                                    dst.Tables("SOTORDR5").Rows.Add(rowSOTORDR5)
                                End If

                                rowSOTORDR5.Item("CUST_NAME") = rowARTCUST1.Item("CUST_NAME") & String.Empty
                                rowSOTORDR5.Item("CUST_ADDR1") = rowARTCUST1.Item("CUST_ADDR1") & String.Empty
                                rowSOTORDR5.Item("CUST_ADDR2") = rowARTCUST1.Item("CUST_ADDR2") & String.Empty
                                rowSOTORDR5.Item("CUST_CITY") = rowARTCUST1.Item("CUST_CITY") & String.Empty
                                rowSOTORDR5.Item("CUST_STATE") = rowARTCUST1.Item("CUST_STATE") & String.Empty
                                rowSOTORDR5.Item("CUST_ZIP_CODE") = rowARTCUST1.Item("CUST_ZIP_CODE") & String.Empty
                                rowSOTORDR5.Item("CUST_PHONE") = rowARTCUST1.Item("CUST_PHONE") & String.Empty
                                rowSOTORDR5.Item("CUST_COUNTRY") = rowARTCUST1.Item("CUST_COUNTRY") & String.Empty
                                rowSOTORDR5.Item("CUST_FAX") = rowARTCUST1.Item("CUST_FAX") & String.Empty
                                rowSOTORDR5.Item("CUST_EMAIL") = rowARTCUST1.Item("CUST_EMAIL") & String.Empty
                            End If

                            If CUST_CODE = "014211" Then 'FOR NOW WE USE THINNING PRISM FOR EVERYONE EXCEPT FOR DR BEN NAYOR
                                rowDETJOBM1.Item("USE_THINNING_PRISM") = "0"
                            End If

                            rowDETJOBM1.Item("CUST_SHIP_TO_NO") = CUST_SHIP_TO_NO
                            If rowARTCUST2 IsNot Nothing Then
                                rowDETJOBM1.Item("CUST_NAME") = rowARTCUST2.Item("CUST_SHIP_TO_NAME") & String.Empty

                                ' Update Sotordr5 ST Record
                                If dst.Tables("SOTORDR5").Select("ORDR_NO = '" & ORDR_NO & "' and CUST_ADDR_TYPE = 'ST'").Length > 0 Then
                                    rowSOTORDR5 = dst.Tables("SOTORDR5").Select("ORDR_NO = '" & ORDR_NO & "' and CUST_ADDR_TYPE = 'ST'")(0)
                                Else
                                    rowSOTORDR5 = dst.Tables("SOTORDR5").NewRow
                                    rowSOTORDR5.Item("ORDR_NO") = ORDR_NO
                                    rowSOTORDR5.Item("CUST_ADDR_TYPE") = "ST"
                                    dst.Tables("SOTORDR5").Rows.Add(rowSOTORDR5)
                                End If

                                rowSOTORDR5.Item("CUST_NAME") = rowARTCUST2.Item("CUST_SHIP_TO_NAME") & String.Empty
                                rowSOTORDR5.Item("CUST_ADDR1") = rowARTCUST2.Item("CUST_SHIP_TO_ADDR1") & String.Empty
                                rowSOTORDR5.Item("CUST_ADDR2") = rowARTCUST2.Item("CUST_SHIP_TO_ADDR2") & String.Empty
                                rowSOTORDR5.Item("CUST_CITY") = rowARTCUST2.Item("CUST_SHIP_TO_CITY") & String.Empty
                                rowSOTORDR5.Item("CUST_STATE") = rowARTCUST2.Item("CUST_SHIP_TO_STATE") & String.Empty
                                rowSOTORDR5.Item("CUST_ZIP_CODE") = rowARTCUST2.Item("CUST_SHIP_TO_ZIP_CODE") & String.Empty
                                rowSOTORDR5.Item("CUST_PHONE") = rowARTCUST2.Item("CUST_SHIP_TO_PHONE") & String.Empty
                                rowSOTORDR5.Item("CUST_COUNTRY") = rowARTCUST2.Item("CUST_SHIP_TO_COUNTRY") & String.Empty
                                rowSOTORDR5.Item("CUST_FAX") = rowARTCUST2.Item("CUST_SHIP_TO_FAX") & String.Empty
                                rowSOTORDR5.Item("CUST_EMAIL") = rowARTCUST2.Item("CUST_SHIP_TO_EMAIL") & String.Empty
                            ElseIf rowARTCUST1 IsNot Nothing Then
                                ' Update Sotordr5 ST Record
                                If dst.Tables("SOTORDR5").Select("ORDR_NO = '" & ORDR_NO & "' and CUST_ADDR_TYPE = 'ST'").Length > 0 Then
                                    rowSOTORDR5 = dst.Tables("SOTORDR5").Select("ORDR_NO = '" & ORDR_NO & "' and CUST_ADDR_TYPE = 'ST'")(0)
                                Else
                                    rowSOTORDR5 = dst.Tables("SOTORDR5").NewRow
                                    rowSOTORDR5.Item("ORDR_NO") = ORDR_NO
                                    rowSOTORDR5.Item("CUST_ADDR_TYPE") = "ST"
                                    dst.Tables("SOTORDR5").Rows.Add(rowSOTORDR5)
                                End If

                                rowSOTORDR5.Item("CUST_NAME") = rowARTCUST1.Item("CUST_NAME") & String.Empty
                                rowSOTORDR5.Item("CUST_ADDR1") = rowARTCUST1.Item("CUST_ADDR1") & String.Empty
                                rowSOTORDR5.Item("CUST_ADDR2") = rowARTCUST1.Item("CUST_ADDR2") & String.Empty
                                rowSOTORDR5.Item("CUST_CITY") = rowARTCUST1.Item("CUST_CITY") & String.Empty
                                rowSOTORDR5.Item("CUST_STATE") = rowARTCUST1.Item("CUST_STATE") & String.Empty
                                rowSOTORDR5.Item("CUST_ZIP_CODE") = rowARTCUST1.Item("CUST_ZIP_CODE") & String.Empty
                                rowSOTORDR5.Item("CUST_PHONE") = rowARTCUST1.Item("CUST_PHONE") & String.Empty
                                rowSOTORDR5.Item("CUST_COUNTRY") = rowARTCUST1.Item("CUST_COUNTRY") & String.Empty
                                rowSOTORDR5.Item("CUST_FAX") = rowARTCUST1.Item("CUST_FAX") & String.Empty
                                rowSOTORDR5.Item("CUST_EMAIL") = rowARTCUST1.Item("CUST_EMAIL") & String.Empty

                            End If

                            ' Need to update LENS_DESIGNER_CODE
                            Dim rowDETDSGN1 As DataRow = Nothing
                            If rowDETJOBM1.Item("LENS_DESIGNER_CODE") & String.Empty = String.Empty Then
                                rowDETDSGN1 = baseClass.LookUp("DETDSGN1", rowDETJOBM1.Item("LENS_DESIGN_CODE"))
                                If rowDETDSGN1 IsNot Nothing Then
                                    rowDETJOBM1.Item("LENS_DESIGNER_CODE") = rowDETDSGN1.Item("LENS_DESIGNER_CODE") & String.Empty
                                End If
                            End If

                            ' Set Enter As Worn defaults
                            If rowDETDSGN1 IsNot Nothing AndAlso rowDETDSGN1.Item("SHOW_AS_WORN") & String.Empty = "1" Then
                                For Each VCA_KEY As String In New String() {"ZTILT", "PANTO", "BVD", "RVD"}
                                    Dim row As DataRow = baseClass.LookUp("DETDEFD1", New String() {rowDETJOBM1.Item("LENS_DESIGN_CODE"), VCA_KEY})
                                    If row IsNot Nothing Then
                                        Dim VALUE_NUM As Decimal = Val(row.Item("VALUE_NUM") & String.Empty)
                                        Dim COLUMN_NAME As String = ""
                                        If VCA_KEY = "BVD" Then
                                            COLUMN_NAME = "FITTING_VERTEX"
                                        ElseIf VCA_KEY = "PANTO" Then
                                            COLUMN_NAME = "PANTOSCOPIC_TILT"
                                        ElseIf VCA_KEY = "RVD" Then
                                            COLUMN_NAME = "REFRACTIVE_VERTEX"
                                        ElseIf VCA_KEY = "ZTILT" Then
                                            COLUMN_NAME = "PANORAMIC_ANGLE"
                                        End If
                                        dst.Tables("DETJOBM3").Select("JOB_NO = '" & JOB_NO & "' AND RL = 'R'")(0).Item(COLUMN_NAME) = VALUE_NUM
                                        dst.Tables("DETJOBM3").Select("JOB_NO = '" & JOB_NO & "' AND RL = 'L'")(0).Item(COLUMN_NAME) = VALUE_NUM
                                    End If
                                Next
                            End If

                            ' Now see if they provided values for 'Enter As Worn' and overwrite the defaults
                            ' All items in the Select Case statement were emailed to Maria and Dana on 10/17/2012
                            ' I asked what to do with these values
                            For Each rowPERSONALIZED_DATA As DataRow In vwXmlDataset.Tables("PERSONALIZED_DATA").Select("PATIENT_ID = " & PATIENT_ID)
                                Dim PERSONALIZED_DATA_ID As String = rowPERSONALIZED_DATA.Item("PERSONALIZED_DATA_ID") & String.Empty
                                Dim COLUMN_NAME As String = String.Empty
                                Dim eye As String = String.Empty
                                For Each rowSPECIAL_PARAMETER As DataRow In vwXmlDataset.Tables("SPECIAL_PARAMETER").Select("PERSONALIZED_DATA_ID = " & PERSONALIZED_DATA_ID)
                                    COLUMN_NAME = String.Empty
                                    eye = String.Empty
                                    Select Case rowSPECIAL_PARAMETER.Item("NAME") & String.Empty
                                        Case "HE_coeff"
                                        Case "ST_coeff"
                                        Case "Progression_Length"
                                        Case "VertexDistance"
                                        Case "WrapAngle" : COLUMN_NAME = "PANORAMIC_ANGLE" : eye = "B"
                                        Case "PantoAngle" : COLUMN_NAME = "PANTOSCOPIC_TILT" : eye = "B"
                                        Case "SeeProudRightInsert"
                                        Case "SeeProudLeftInsert"
                                        Case "RightFittingHeight"
                                        Case "LeftFittingHeight"
                                        Case "RightVertexDistance" : COLUMN_NAME = "FITTING_VERTEX" : eye = "R"
                                        Case "LeftVertexDistance" : COLUMN_NAME = "FITTING_VERTEX" : eye = "L"
                                        Case "ReadingDistance"
                                        Case "CorridorLength"
                                        Case "FrameFit"
                                        Case "RightERCD"
                                        Case "LeftERCD"
                                        Case "CAPE"
                                    End Select

                                    If COLUMN_NAME.Length > 0 Then
                                        If (rowSPECIAL_PARAMETER.Item("VALUE") & String.Empty).ToString.Length > 0 Then
                                            Dim VALUE_NUM As Decimal = Val(rowSPECIAL_PARAMETER.Item("VALUE") & String.Empty)
                                            If eye = "B" Then
                                                dst.Tables("DETJOBM3").Select("JOB_NO = '" & JOB_NO & "' AND RL = 'R'")(0).Item(COLUMN_NAME) = VALUE_NUM
                                                dst.Tables("DETJOBM3").Select("JOB_NO = '" & JOB_NO & "' AND RL = 'L'")(0).Item(COLUMN_NAME) = VALUE_NUM
                                            ElseIf eye = "L" Then
                                                dst.Tables("DETJOBM3").Select("JOB_NO = '" & JOB_NO & "' AND RL = 'L'")(0).Item(COLUMN_NAME) = VALUE_NUM
                                            ElseIf eye = "R" Then
                                                dst.Tables("DETJOBM3").Select("JOB_NO = '" & JOB_NO & "' AND RL = 'R'")(0).Item(COLUMN_NAME) = VALUE_NUM
                                            End If
                                        End If
                                    End If
                                Next
                            Next

                            ' Other processing for LENS_DESIGNER_CODE
                            Dim rowDESIGN0 As DataRow = baseClass.LookUp("DETDSGN0", rowDETJOBM1.Item("LENS_DESIGNER_CODE") & String.Empty)
                            If rowDESIGN0 IsNot Nothing Then
                                Dim BLANK_SELECTION_VIA_LDS As String = rowDESIGN0.Item("BLANK_SELECTION_VIA_LDS") & String.Empty
                                If BLANK_SELECTION_VIA_LDS = "1" Then
                                    If rowDETDSGN1.Item("CUST_SUPPLIED_BLANKS") & "" <> "1" Then
                                        rowDETJOBM1.Item("LDS_IND") = "0"
                                        rowDETJOBM1.Item("LDS_QUEUED") = DBNull.Value
                                        If rowDETDSGN1.Item("LENS_DESIGN_SLCT_BLNK") & "" = "1" Then
                                            rowDETJOBM1.Item("BLANK_SELECTION") = "M"
                                        Else
                                            rowDETJOBM1.Item("BLANK_SELECTION") = "D"
                                        End If
                                    End If
                                End If
                            End If

                            If rowARTCUST1 IsNot Nothing AndAlso rowARTCUST1.Item("CUST_DEL_JOB_INSPCT_SUP") & String.Empty <> String.Empty Then
                                rowDETJOBM1.Item("JOB_INSPCT_SUP") = rowARTCUST1.Item("CUST_DEL_JOB_INSPCT_SUP") & String.Empty
                            End If

                            ' Trying this so I can evalaute data in the routine
                            errors = ValidateJobs(rowDETJOBM1)

                            If Val(rowDETJOBM1.Item("FRAME_A_WIDTH") & "") > Val(dst.Tables("DETPARM1").Rows(0).Item("DE_PARM_MAX_A") & "") _
                                OrElse Val(rowDETJOBM1.Item("FRAME_B_HEIGHT") & "") > Val(dst.Tables("DETPARM1").Rows(0).Item("DE_PARM_MAX_B") & "") _
                                OrElse Val(rowDETJOBM1.Item("FRAME_DBL_BRIDGE") & "") > Val(dst.Tables("DETPARM1").Rows(0).Item("DE_PARM_MAX_DBL") & "") _
                                OrElse Val(rowDETJOBM1.Item("FRAME_ED_DIAGONAL") & "") > Val(dst.Tables("DETPARM1").Rows(0).Item("DE_PARM_MAX_ED") & "") _
                                OrElse Val(rowDETJOBM1.Item("FRAME_A_WIDTH") & "") < 0 _
                                OrElse Val(rowDETJOBM1.Item("FRAME_B_HEIGHT") & "") < 0 _
                                OrElse Val(rowDETJOBM1.Item("FRAME_DBL_BRIDGE") & "") < 0 _
                                OrElse Val(rowDETJOBM1.Item("FRAME_ED_DIAGONAL") & "") < 0 Then

                                Dim EMsg As String = "Max Values Frame Measurements: A=" & CStr(Val(dst.Tables("DETPARM1").Rows(0).Item("DE_PARM_MAX_A") & "")) _
                                & ", B=" & CStr(Val(dst.Tables("DETPARM1").Rows(0).Item("DE_PARM_MAX_B") & "")) & ", ED=" _
                                & CStr(Val(dst.Tables("DETPARM1").Rows(0).Item("DE_PARM_MAX_ED") & "")) & ", DBL=" _
                                & CStr(Val(dst.Tables("DETPARM1").Rows(0).Item("DE_PARM_MAX_DBL") & "")) & ""

                                errors.Add(New DelJobService.DelJobValidationError("Customer", dst.Tables("SOTORDRO").Select("ORDR_REL_HOLD_CODES = '" & code & "'")(0).Item("ORDR_COMMENT")))

                            End If


                            Dim rowSOTORDR1 As DataRow = dst.Tables("SOTORDR1").NewRow
                            rowSOTORDR1.Item("ORDR_NO") = ORDR_NO
                            rowSOTORDR1.Item("ORDR_DATE") = rowDETJOBM1.Item("ORDR_DATE")
                            rowSOTORDR1.Item("CUST_CODE") = CUST_CODE
                            rowSOTORDR1.Item("CUST_NAME") = rowDETJOBM1.Item("CUST_NAME") & String.Empty
                            rowSOTORDR1.Item("CUST_SHIP_TO_NO") = rowDETJOBM1.Item("CUST_SHIP_TO_NO") & String.Empty
                            rowSOTORDR1.Item("CUST_SHIP_TO_NAME") = rowDETJOBM1.Item("CUST_SHIP_TO_NAME") & String.Empty
                            rowSOTORDR1.Item("ORDR_CUST_PO") = rowDETJOBM1.Item("ORDR_CUST_PO") & String.Empty
                            rowSOTORDR1.Item("ORDR_STATUS") = "O"
                            rowSOTORDR1.Item("ORDR_SOURCE") = ORDR_SOURCE
                            'rowSOTORDR1.Item("OPS_YYYYPP") = ABSolution.ASCMAIN1.CYP
                            rowSOTORDR1.Item("SHIP_VIA_CODE") = rowDETJOBM1.Item("SHIP_VIA_CODE") & String.Empty
                            dst.Tables("SOTORDR1").Rows.Add(rowSOTORDR1)

                            Dim errorCode As String = String.Empty
                            SetBillToAttributes(CUST_CODE, CUST_SHIP_TO_NO, rowSOTORDR1, errorCode)
                            CreateSalesOrderTax(ORDR_NO)

                            For Each errCode As Char In errorCode
                                If dst.Tables("SOTORDRO").Select("ORDR_REL_HOLD_CODES = '" & errCode & "'").Length > 0 Then
                                    errors.Add(New DelJobService.DelJobValidationError("Customer", dst.Tables("SOTORDRO").Select("ORDR_REL_HOLD_CODES = '" & errCode & "'")(0).Item("ORDR_COMMENT")))
                                End If
                            Next

                            rowSOTORDR1.Item("WHSE_CODE") = "003"
                            rowSOTORDR1.Item("INIT_OPER") = ABSolution.ASCMAIN1.USER_ID
                            rowSOTORDR1.Item("LAST_OPER") = ABSolution.ASCMAIN1.USER_ID
                            rowSOTORDR1.Item("INIT_DATE") = DateTime.Now
                            rowSOTORDR1.Item("LAST_DATE") = DateTime.Now
                            rowSOTORDR1.Item("ORDR_TYPE_CODE") = "REG"
                            rowSOTORDR1.Item("ORDR_CALLER_NAME") = "VW Import"
                            rowSOTORDR1.Item("ORDR_DPD") = "0"
                            rowSOTORDR1.Item("BRANCH_CODE") = "NY"
                            rowSOTORDR1.Item("DIVISION_CODE") = "DEL"
                            rowSOTORDR1.Item("ORDR_TOTAL_AMT") = 0
                            rowSOTORDR1.Item("ORDR_SALES") = 0

                            ERR_LNO = 1
                            For Each err As DelJobService.DelJobValidationError In errors
                                Dim rowDETJOBIE As DataRow = dst.Tables("DETJOBIE").NewRow
                                rowDETJOBIE.Item("JOB_NO") = JOB_NO
                                rowDETJOBIE.Item("ERR_LNO") = ERR_LNO
                                ERR_LNO += 1
                                rowDETJOBIE.Item("ERR_CODE") = err.Field
                                rowDETJOBIE.Item("ERR_DESC") = TruncateField(err.Message, "DETJOBIE", "ERR_DESC")
                                dst.Tables("DETJOBIE").Rows.Add(rowDETJOBIE)
                            Next

                            PriceDelJob(JOB_NO)

                            rowDETJOBM1.Item("TERM_CODE") = rowSOTORDR1.Item("TERM_CODE") & String.Empty

                            rowDETDSGN1 = baseClass.LookUp("DETDSGN1", New String() {rowDETJOBM1.Item("LENS_DESIGNER_CODE") & String.Empty})
                            If rowDETDSGN1 IsNot Nothing AndAlso rowDETDSGN1.Item("LENS_TYPE") & String.Empty = "S" Then
                                For Each rowDETJOBM3 As DataRow In dst.Tables("DETJOBM3").Rows
                                    rowDETJOBM3.Item("ADD_POWER") = DBNull.Value
                                Next
                            End If

                            Dim rowDETDSGN3 As DataRow = baseClass.LookUp("DETDSGN3", New String() _
                                    {rowDETJOBM1.Item("LENS_DESIGN_CODE") & String.Empty, _
                                     rowDETJOBM1.Item("MATL_CODE") & String.Empty, _
                                     rowDETJOBM1.Item("COLOR_CODE") & String.Empty})

                            If rowDETDSGN3 IsNot Nothing Then
                                rowDETJOBM1.Item("LIST_PRICE") = rowDETDSGN3("LIST_PRICE")
                            End If

                            Dim INV_SALES As Decimal = Val(dst.Tables("DETJOBM2").Compute("SUM(JOB_AMT)", String.Empty) & String.Empty)
                            rowDETJOBM1.Item("INV_SALES") = INV_SALES

                            If errors.Count > 0 Then
                                rowDETJOBM1.Item("JOB_STATUS") = "H"
                                ' If on hold then take out of queue
                                rowDETJOBM1.Item("JOB_IN_QUEUE") = "0"
                                rowSOTORDR1.Item("ORDR_STATUS") = "H"
                            End If

                            Dim rowDETJOBM4 As DataRow = dst.Tables("DETJOBM4").NewRow
                            rowDETJOBM4.Item("JOB_NO") = JOB_NO
                            rowDETJOBM4.Item("STATUS_CODE") = "110"
                            rowDETJOBM4.Item("INIT_OPER") = rowDETJOBM1.Item("INIT_OPER")
                            rowDETJOBM4.Item("INIT_DATE") = rowDETJOBM1.Item("INIT_DATE")
                            dst.Tables("DETJOBM4").Rows.Add(rowDETJOBM4)

                            With baseClass
                                .BeginTrans()
                                inTrans = True

                                .clsASCBASE1.Update_Record_TDA("DETJOBM1")
                                .clsASCBASE1.Update_Record_TDA("DETJOBM2")
                                .clsASCBASE1.Update_Record_TDA("DETJOBM3")
                                .clsASCBASE1.Update_Record_TDA("DETJOBM4")
                                .clsASCBASE1.Update_Record_TDA("DETJOBIE")

                                .clsASCBASE1.Update_Record_TDA("SOTORDR1")
                                .clsASCBASE1.Update_Record_TDA("SOTORDR5")

                                If dst.Tables("DETJOBIE").Rows.Count = 0 Then
                                    ABSolution.ASCDATA1.ExecuteSQL("Begin ARPCUST6_ORDR_NO('DEL', '" & rowDETJOBM1.Item("CUST_CODE") & "', '" & rowDETJOBM1.Item("JOB_NO") & "', SYSDATE, " & INV_SALES & ");  End;")
                                End If

                                .CommitTrans()
                                inTrans = False
                            End With
                        Next ' For Each rowRX_SPECTACLE

                    Catch ex As Exception
                        RecordLogEntry("ProcessVisionWebDELOrders: File(" & orderFile & ") " & ex.Message)
                        Dim rowDETJOBIE As DataRow = dst.Tables("DETJOBIE").NewRow
                        rowDETJOBIE.Item("JOB_NO") = JOB_NO
                        rowDETJOBIE.Item("ERR_LNO") = ERR_LNO
                        ERR_LNO += 1
                        rowDETJOBIE.Item("ERR_CODE") = "System Error"
                        rowDETJOBIE.Item("ERR_DESC") = TruncateField(ex.Message, "DETJOBIE", "ERR_DESC")
                        dst.Tables("DETJOBIE").Rows.Add(rowDETJOBIE)
                    Finally
                        If Not invalidCustomer Then
                            My.Computer.FileSystem.MoveFile(orderFile, vwConnection.LocalInDirArchive & My.Computer.FileSystem.GetName(orderFile), True)
                            If dst.Tables("DETJOBIE").Rows.Count > 0 Then
                                Dim note As String = "Order File: " & orderFile & Environment.NewLine & Environment.NewLine
                                For Each row As DataRow In dst.Tables("DETJOBIE").Select
                                    note &= Environment.NewLine & row.Item("ERR_CODE") & vbTab & row.Item("ERR_DESC") & Environment.NewLine
                                Next
                                emailErrors(String.Empty, dst.Tables("DETJOBIE").Rows.Count, note, "DEL_VWEB")
                                ImportErrorNotification.Clear()
                            End If
                        End If
                        numJobsProcessed += 1
                    End Try
                Next ' For Each orderFile

            Catch ex As Exception
                RecordLogEntry("ProcessVisionWebDELOrders: " & ex.Message)
                Dim note As String = "ProcessVisionWebDELOrders: " & ex.Message & Environment.NewLine
                emailErrors(String.Empty, 1, note, "DEL_VWEB")
            Finally
                RecordLogEntry("ProcessVisionWebDELOrders: " & numJobsProcessed & " DEL orders imported ")
                ' Archive processed XML files
            End Try
        End Sub

        Private Function ValidateJobs(ByRef rowDETJOBM1 As DataRow) As List(Of DelJobService.DelJobValidationError)
            Dim errors As List(Of DelJobService.DelJobValidationError) = New List(Of DelJobService.DelJobValidationError)

            Try
                Dim jb As DelJobService = New DelJobService(ABSolution.ASCMAIN1.oraCon, ABSolution.ASCMAIN1.oraAda)

                ' JOB Status 110 = Entered
                Dim job As DelJobService.JobValidatorFields = jb.CreateJobObject(rowDETJOBM1, dst.Tables("DETJOBM3"), "110")
                errors = jb.ValidateJobData(job)

            Catch ex As Exception
                errors.Add(New DelJobService.DelJobValidationError("ValidateJobs", ex.Message))
            End Try

            Return errors

        End Function

        Private Sub PriceDelJob(ByVal JOB_NO As String)

            'MAKE EACH CALL A BOOLEAN AND RECORD AN ERROR CODE IF ERROR PRICING
            Try
                LoadPricingLines(JOB_NO)
                Charge_Lens(JOB_NO)
                Charge_Balance_Lens(JOB_NO)
                Charge_Fog_Free(JOB_NO)
                Charge_Coating(JOB_NO)
                Charge_Mirror_Coating(JOB_NO)
                Charge_WrapEdge(JOB_NO)
                Charge_Tinting(JOB_NO)
                Charge_Edging(JOB_NO)
                Charge_Polishing(JOB_NO)

            Catch ex As Exception
                RecordLogEntry("PriceDelJob: (" & JOB_NO & ") - " & ex.Message)
            End Try

        End Sub

        Private Sub LoadPricingLines(ByVal JOB_NO As String)

            Try
                'Create price line in detjomb2
                If dst.Tables("DETJOBM2").Rows.Find(New Object() {JOB_NO, 11}) Is Nothing Then
                    dst.Tables("DETJOBM2").Rows.Add(New Object() {JOB_NO, 11, "L", "R", 0, 0, 0})
                End If

                If dst.Tables("DETJOBM2").Rows.Find(New Object() {JOB_NO, 12}) Is Nothing Then
                    dst.Tables("DETJOBM2").Rows.Add(New Object() {JOB_NO, 12, "L", "L", 0, 0, 0})
                End If

                If dst.Tables("DETJOBM2").Rows.Find(New Object() {JOB_NO, 20}) Is Nothing Then
                    dst.Tables("DETJOBM2").Rows.Add(New Object() {JOB_NO, 20, "C", "", 0, 0, 0})
                End If
                If dst.Tables("DETJOBM2").Rows.Find(New Object() {JOB_NO, 22}) Is Nothing Then
                    dst.Tables("DETJOBM2").Rows.Add(New Object() {JOB_NO, 22, "M", "", 0, 0, 0})
                End If

                If dst.Tables("DETJOBM2").Rows.Find(New Object() {JOB_NO, 30}) Is Nothing Then
                    dst.Tables("DETJOBM2").Rows.Add(New Object() {JOB_NO, 30, "T", "", 0, 0, 0})
                End If

                If dst.Tables("DETJOBM2").Rows.Find(New Object() {JOB_NO, 40}) Is Nothing Then
                    dst.Tables("DETJOBM2").Rows.Add(New Object() {JOB_NO, 40, "E", "", 0, 0, 0})
                End If

                If dst.Tables("DETJOBM2").Rows.Find(New Object() {JOB_NO, 50}) Is Nothing Then
                    dst.Tables("DETJOBM2").Rows.Add(New Object() {JOB_NO, 50, "P", "", 0, 0, 0})
                End If

                If dst.Tables("DETJOBM2").Rows.Find(New Object() {JOB_NO, 60}) Is Nothing Then
                    dst.Tables("DETJOBM2").Rows.Add(New Object() {JOB_NO, 60, "B", "", 0, 0, 0})
                End If

                If dst.Tables("DETJOBM2").Rows.Find(New Object() {JOB_NO, 70}) Is Nothing Then
                    dst.Tables("DETJOBM2").Rows.Add(New Object() {JOB_NO, 70, "W", "", 0, 0, 0})
                End If

                If dst.Tables("DETJOBM2").Rows.Find(New Object() {JOB_NO, 80}) Is Nothing Then
                    dst.Tables("DETJOBM2").Rows.Add(New Object() {JOB_NO, 80, "F", "", 0, 0, 0})
                End If

            Catch ex As Exception
                RecordLogEntry("LoadPricingLines: " & ex.Message)
            End Try
        End Sub

        Private Sub Charge_Coating(ByVal JOB_NO As String)

            Try

                If dst.Tables("DETJOBM1").Select("JOB_NO = '" & JOB_NO & "'").Length = 0 Then
                    Exit Sub
                End If

                If dst.Tables("DETJOBM2").Select("JOB_NO = '" & JOB_NO & "'").Length = 0 Then
                    Exit Sub
                End If

                If dst.Tables("DETJOBM2").Select("JOB_NO = '" & JOB_NO & "' AND JOB_CHARGE_TYPE = 'C'", "").Length = 0 Then
                    Exit Sub
                End If

                Dim rowDETJOBM1 As DataRow = dst.Tables("DETJOBM1").Select("JOB_NO = '" & JOB_NO & "'")(0)
                Dim rowDETJOBM2 As DataRow = dst.Tables("DETJOBM2").Select("JOB_NO = '" & JOB_NO & "' AND JOB_CHARGE_TYPE = 'C'", "")(0)

                If rowDETJOBM1.Item("AR_COATING") & String.Empty = "1" Then
                    rowDETJOBM2("JOB_QTY") = IIf(rowDETJOBM1.Item("LENS_ORDER") & String.Empty = "B" Or rowDETJOBM1.Item("BALANCE_LENS") & String.Empty = "1", 2, 1)
                Else
                    rowDETJOBM2("JOB_QTY") = 0
                End If

                Dim LIST_PRICE As Decimal = Val(dst.Tables("DETPARM1").Rows(0).Item("DE_PARM_COATING") & String.Empty) / 2
                Dim CUST_PRICE As Decimal = LIST_PRICE

                ' See if the customer has special pricing.
                Dim sql As String = String.Empty
                sql = "SELECT * FROM DETDSGND WHERE CUST_CODE = :PARM1 AND LENS_DESIGN_CODE = :PARM2 AND SERVICE_CODE = 'C'"
                Dim rowDETDSGND As DataRow = ABSolution.ASCDATA1.GetDataRow(sql, "VV", New Object() {rowDETJOBM1.Item("CUST_CODE") & String.Empty, rowDETJOBM1.Item("LENS_DESIGN_CODE") & String.Empty})

                sql = "SELECT * FROM DETCUSTP WHERE CUST_CODE = :PARM1"
                Dim rowDETCUSTP As DataRow = ABSolution.ASCDATA1.GetDataRow(sql, "V", New Object() {rowDETJOBM1.Item("CUST_CODE") & String.Empty})

                If rowDETDSGND IsNot Nothing Then
                    CUST_PRICE = Val(rowDETDSGND.Item("CUST_PRICE") & String.Empty) / 2
                ElseIf rowDETCUSTP IsNot Nothing AndAlso rowDETCUSTP.Item("CUST_COATING_OVERRIDE") & String.Empty = "1" Then
                    CUST_PRICE = Val(rowDETCUSTP.Item("CUST_COATING_PRICE") & String.Empty) / 2
                End If

                rowDETJOBM2("LIST_PRICE") = LIST_PRICE

                Dim INV_DISC_PCT As Decimal = 0
                Dim JOB_PRICE As Decimal = CUST_PRICE * (100 - INV_DISC_PCT) / 100
                rowDETJOBM2("JOB_PRICE") = JOB_PRICE
                rowDETJOBM2("CUST_PRICE") = CUST_PRICE

            Catch ex As Exception
                RecordLogEntry("Charge_Coating: (" & JOB_NO & ") - " & ex.Message)
            End Try

        End Sub

        Private Sub Charge_Edging(ByVal JOB_NO As String)

            Try
                If dst.Tables("DETJOBM1").Select("JOB_NO = '" & JOB_NO & "'").Length = 0 Then
                    Exit Sub
                End If

                If dst.Tables("DETJOBM2").Select("JOB_NO = '" & JOB_NO & "'").Length = 0 Then
                    Exit Sub
                End If

                If dst.Tables("DETJOBM2").Select("JOB_NO = '" & JOB_NO & "' AND JOB_CHARGE_TYPE = 'E'", "").Length = 0 Then
                    Exit Sub
                End If

                Dim rowDETJOBM1 As DataRow = dst.Tables("DETJOBM1").Select("JOB_NO = '" & JOB_NO & "'")(0)
                Dim rowDETJOBM2 As DataRow = dst.Tables("DETJOBM2").Select("JOB_NO = '" & JOB_NO & "' AND JOB_CHARGE_TYPE = 'E'", "")(0)


                If rowDETJOBM1.Item("FINISHED") & String.Empty <> "U" And rowDETJOBM1.Item("EDGING") & String.Empty = "1" Then
                    rowDETJOBM2("JOB_QTY") = IIf(rowDETJOBM1.Item("LENS_ORDER") & String.Empty = "B" Or rowDETJOBM1.Item("BALANCE_LENS") & String.Empty = "1", 2, 1)
                Else
                    rowDETJOBM2("JOB_QTY") = 0
                End If


                Dim LIST_PRICE As Decimal = 0
                Dim CUST_PRICE As Decimal = 0

                Dim sql As String = "Select EDGING_SVC_PRICE from DETFRAM2 " _
                    & " where FRAME_TYPE_CODE = :PARM1" _
                    & " and MATL_LMATTYPE = (Select MATL_LMATTYPE from DETMATL1 where MATL_CODE = :PARM2)"

                ' Validate that Edging is a valid charge
                Dim rowValue As DataRow = ABSolution.ASCDATA1.GetDataRow(sql, "VV", New Object() {rowDETJOBM1.Item("FRAME_TYPE_CODE") & String.Empty, rowDETJOBM1.Item("MATL_CODE") & String.Empty})

                ' See if the customer has special pricing.
                If rowValue IsNot Nothing Then
                    LIST_PRICE = Val(rowValue.Item("EDGING_SVC_PRICE") & String.Empty) / 2
                    CUST_PRICE = LIST_PRICE

                    sql = "SELECT * FROM DETDSGND WHERE CUST_CODE = :PARM1 AND LENS_DESIGN_CODE = :PARM2 AND SERVICE_CODE = 'E'"
                    Dim rowDETDSGND As DataRow = ABSolution.ASCDATA1.GetDataRow(sql, "VV", New Object() {rowDETJOBM1.Item("CUST_CODE") & String.Empty, rowDETJOBM1.Item("LENS_DESIGN_CODE") & String.Empty})

                    sql = "SELECT * FROM DETCUSTP WHERE CUST_CODE = :PARM1"
                    Dim rowDETCUSTP As DataRow = ABSolution.ASCDATA1.GetDataRow(sql, "V", New Object() {rowDETJOBM1.Item("CUST_CODE") & String.Empty})

                    If rowDETDSGND IsNot Nothing Then
                        CUST_PRICE = Val(rowDETDSGND.Item("CUST_PRICE") & String.Empty) / 2
                    ElseIf rowDETCUSTP IsNot Nothing AndAlso rowDETCUSTP.Item("CUST_EDGING_OVERRIDE") & String.Empty = "1" Then
                        CUST_PRICE = Val(rowDETCUSTP.Item("CUST_EDGING_PRICE") & String.Empty) / 2
                    End If
                End If

                rowDETJOBM2("LIST_PRICE") = LIST_PRICE

                Dim INV_DISC_PCT As Decimal = 0
                Dim JOB_PRICE As Decimal = CUST_PRICE * (100 - INV_DISC_PCT) / 100
                rowDETJOBM2("JOB_PRICE") = JOB_PRICE
                rowDETJOBM2("CUST_PRICE") = CUST_PRICE

            Catch ex As Exception
                RecordLogEntry("Charge_Edging: (" & JOB_NO & ") - " & ex.Message)
            End Try
        End Sub

        Private Sub Charge_Lens(ByVal JOB_NO As String)

            Try
                If dst.Tables("DETJOBM1").Select("JOB_NO = '" & JOB_NO & "'").Length = 0 Then
                    Exit Sub
                End If

                If dst.Tables("DETJOBM2").Select("JOB_NO = '" & JOB_NO & "'").Length = 0 Then
                    Exit Sub
                End If

                Dim rowDETJOBM1 As DataRow = dst.Tables("DETJOBM1").Select("JOB_NO = '" & JOB_NO & "'")(0)
                Dim rowDETJOBM2 As DataRow = dst.Tables("DETJOBM2").Select("JOB_NO = '" & JOB_NO & "'")(0)
                Dim rowDETJOBM2_R As DataRow = dst.Tables("DETJOBM2").Select("JOB_NO = '" & JOB_NO & "' AND JOB_CHARGE_TYPE = 'L' and JOB_CHARGE_EYE = 'R'", "")(0)
                Dim rowDETJOBM2_L As DataRow = dst.Tables("DETJOBM2").Select("JOB_NO = '" & JOB_NO & "' AND JOB_CHARGE_TYPE = 'L' and JOB_CHARGE_EYE = 'L'", "")(0)

                Dim LENS_ORDER As String = rowDETJOBM1.Item("LENS_ORDER") & String.Empty
                Dim FINISHED As String = rowDETJOBM1.Item("FINISHED") & String.Empty

                rowDETJOBM2_R("JOB_QTY") = IIf(LENS_ORDER = "L" Or FINISHED = "O", 0, 1)
                rowDETJOBM2_L("JOB_QTY") = IIf(LENS_ORDER = "R" Or FINISHED = "O", 0, 1)


                Dim LIST_PRICE As Decimal = 0
                Dim sql As String = "Select LIST_PRICE from DETDSGN3 " _
                    & " where LENS_DESIGN_CODE = '" & rowDETJOBM1.Item("LENS_DESIGN_CODE") & "'" _
                    & " and MATL_CODE = '" & rowDETJOBM1.Item("MATL_CODE") & "'" _
                    & " and COLOR_CODE = '" & rowDETJOBM1.Item("COLOR_CODE") & "'"
                LIST_PRICE = Val(ABSolution.ASCDATA1.GetDataValue(sql)) / 2

                Dim CUST_PRICE As Decimal = LIST_PRICE

                ' See if the customer has special pricing.
                sql = "SELECT * FROM DETDSGNP WHERE CUST_CODE = :PARM1 AND LENS_DESIGN_CODE = :PARM2 AND MATL_CODE = :PARM3 AND COLOR_CODE = :PARM4"
                Dim rowDETDSGNP As DataRow = ABSolution.ASCDATA1.GetDataRow(sql, "VVVV", New String() _
                    {rowDETJOBM1.Item("CUST_CODE"), rowDETJOBM1.Item("LENS_DESIGN_CODE"), rowDETJOBM1.Item("MATL_CODE"), rowDETJOBM1.Item("COLOR_CODE")})

                If rowDETDSGNP IsNot Nothing AndAlso rowDETDSGNP.Item("CUST_PRICE_OVERRIDE") & String.Empty = "1" Then
                    CUST_PRICE = Val(rowDETDSGNP.Item("CUST_PRICE") & String.Empty) / 2
                End If

                Dim INV_DISC_PCT As Decimal = 0
                Dim JOB_PRICE As Decimal = CUST_PRICE * (100 - INV_DISC_PCT) / 100

                rowDETJOBM2_R("CUST_PRICE") = CUST_PRICE
                rowDETJOBM2_L("CUST_PRICE") = CUST_PRICE

                rowDETJOBM2_R("LIST_PRICE") = LIST_PRICE
                rowDETJOBM2_R("JOB_PRICE") = JOB_PRICE

                rowDETJOBM2_L("LIST_PRICE") = LIST_PRICE
                rowDETJOBM2_L("JOB_PRICE") = JOB_PRICE

            Catch ex As Exception
                RecordLogEntry("Charge_Lens" & ex.Message)
            End Try

        End Sub

        Private Sub Charge_Balance_Lens(ByVal JOB_NO As String)

            Dim sql As String = String.Empty

            Try
                If dst.Tables("DETJOBM1").Select("JOB_NO = '" & JOB_NO & "'").Length = 0 Then
                    Exit Sub
                End If

                If dst.Tables("DETJOBM2").Select("JOB_NO = '" & JOB_NO & "'").Length = 0 Then
                    Exit Sub
                End If

                Dim rowDETJOBM1 As DataRow = dst.Tables("DETJOBM1").Select("JOB_NO = '" & JOB_NO & "'")(0)
                Dim rowDETJOBM2 As DataRow = dst.Tables("DETJOBM2").Select("JOB_NO = '" & JOB_NO & "' AND JOB_CHARGE_TYPE = 'B'")(0)
                Dim BALANCE_LENS As String = rowDETJOBM1.Item("BALANCE_LENS") & String.Empty

                If Val(BALANCE_LENS) = 1 Then
                    rowDETJOBM2("JOB_QTY") = 1
                    rowDETJOBM2("JOB_CHARGE_EYE") = IIf(rowDETJOBM1.Item("LENS_ORDER") & String.Empty = "R", "L", "R")
                Else
                    rowDETJOBM2("JOB_QTY") = 0
                    rowDETJOBM2("JOB_CHARGE_EYE") = ""
                End If

                If BALANCE_LENS = "1" OrElse Val(rowDETJOBM2("LIST_PRICE") & "") = 0 Then

                    sql = "SELECT * FROM DETDSGN3 WHERE LENS_DESIGN_CODE = :PARM1 AND MATL_CODE = :PARM2 AND COLOR_CODE = :PARM3"
                    Dim rowDETDSGN3 As DataRow = ABSolution.ASCDATA1.GetDataRow(sql, "VVV", New Object() {rowDETJOBM1.Item("LENS_DESIGN_CODE") & String.Empty, _
                                                                                                          rowDETJOBM1.Item("MATL_CODE") & String.Empty, _
                                                                                                          rowDETJOBM1.Item("COLOR_CODE") & String.Empty})

                    Dim LIST_PRICE As Decimal = 0
                    If rowDETDSGN3 IsNot Nothing Then
                        LIST_PRICE = Val(rowDETDSGN3.Item("BALANCE_PRICE") & String.Empty)
                    End If
                    Dim CUST_PRICE As Decimal = LIST_PRICE

                    ' See if the customer has special pricing.
                    sql = "SELECT * FROM DETDSGNP WHERE CUST_CODE = :PARM1 AND LENS_DESIGN_CODE = :PARM2 AND MATL_CODE = :PARM3 AND COLOR_CODE = :PARM4"
                    Dim rowDETDSGNP As DataRow = ABSolution.ASCDATA1.GetDataRow(sql, "VVVV", New Object() {rowDETJOBM1.Item("CUST_CODE") & String.Empty, _
                                                                                                           rowDETJOBM1.Item("LENS_DESIGN_CODE") & String.Empty, _
                                                                                                           rowDETJOBM1.Item("MATL_CODE") & String.Empty, _
                                                                                                           rowDETJOBM1.Item("COLOR_CODE") & String.Empty})

                    If rowDETDSGNP IsNot Nothing AndAlso rowDETDSGNP.Item("CUST_BALANCE_PRICE_OVERRIDE") & String.Empty = "1" Then
                        CUST_PRICE = Val(rowDETDSGNP.Item("CUST_BALANCE_PRICE") & String.Empty)
                    End If

                    rowDETJOBM2("LIST_PRICE") = LIST_PRICE

                    Dim INV_DISC_PCT As Decimal = 0
                    Dim JOB_PRICE As Decimal = CUST_PRICE * (100 - INV_DISC_PCT) / 100
                    rowDETJOBM2("JOB_PRICE") = JOB_PRICE
                    rowDETJOBM2("CUST_PRICE") = CUST_PRICE
                End If

            Catch ex As Exception
                RecordLogEntry("Charge_Balance_Lens: (" & JOB_NO & ") - " & ex.Message)
            End Try

        End Sub

        Private Sub Charge_Fog_Free(ByVal JOB_NO As String)

            Try

                If dst.Tables("DETJOBM1").Select("JOB_NO = '" & JOB_NO & "'").Length = 0 Then
                    Exit Sub
                End If

                If dst.Tables("DETJOBM2").Select("JOB_NO = '" & JOB_NO & "'").Length = 0 Then
                    Exit Sub
                End If

                If dst.Tables("DETJOBM2").Select("JOB_NO = '" & JOB_NO & "' AND JOB_CHARGE_TYPE = 'F'", "").Length = 0 Then
                    Exit Sub
                End If

                Dim rowDETJOBM1 As DataRow = dst.Tables("DETJOBM1").Select("JOB_NO = '" & JOB_NO & "'")(0)
                Dim rowDETJOBM2 As DataRow = dst.Tables("DETJOBM2").Select("JOB_NO = '" & JOB_NO & "' AND JOB_CHARGE_TYPE = 'F'", "")(0)

                Dim INV_DISC_PCT As Decimal = 0
                Dim JOB_PRICE As Decimal = (DE_PARM_FOG_FREE_COST / 2) * (100 - INV_DISC_PCT) / 100

                Select Case rowDETJOBM1.Item("LENS_ORDER") & String.Empty
                    Case "B"

                    Case Else
                        If Not rowDETJOBM1.Item("BALANCE_LENS") & String.Empty = "1" Then
                            JOB_PRICE = JOB_PRICE / 2
                        End If
                End Select

                JOB_PRICE = JOB_PRICE * (100 - INV_DISC_PCT) / 100

                rowDETJOBM2("CUST_PRICE") = (DE_PARM_FOG_FREE_COST / 2)

                If rowDETJOBM1.Item("FOG_FREE") & String.Empty <> "1" Then
                    rowDETJOBM2.Item("JOB_PRICE") = 0
                    rowDETJOBM2("JOB_QTY") = 0
                Else
                    rowDETJOBM2.Item("JOB_PRICE") = JOB_PRICE
                    rowDETJOBM2("JOB_QTY") = 1
                End If

                If rowDETJOBM1.Item("FOG_FREE") & String.Empty = "1" And Not rowDETJOBM1.Item("MIRROR_COATING") & String.Empty = "1" Then
                    rowDETJOBM1.Item("MIRROR_COATING") = "1"
                    If rowDETJOBM1.Item("MIRROR_COATING_COLOR") & String.Empty = String.Empty Then
                        rowDETJOBM1.Item("MIRROR_COATING_COLOR") = "Fog Free"
                    End If
                End If
            Catch ex As Exception

            End Try

        End Sub

        Private Sub Charge_Mirror_Coating(ByVal JOB_NO As String)

            Try

                If dst.Tables("DETJOBM1").Select("JOB_NO = '" & JOB_NO & "'").Length = 0 Then
                    Exit Sub
                End If

                If dst.Tables("DETJOBM2").Select("JOB_NO = '" & JOB_NO & "'").Length = 0 Then
                    Exit Sub
                End If

                If dst.Tables("DETJOBM2").Select("JOB_NO = '" & JOB_NO & "' AND JOB_CHARGE_TYPE = 'M'", "").Length = 0 Then
                    Exit Sub
                End If

                Dim rowDETJOBM1 As DataRow = dst.Tables("DETJOBM1").Select("JOB_NO = '" & JOB_NO & "'")(0)
                Dim rowDETJOBM2 As DataRow = dst.Tables("DETJOBM2").Select("JOB_NO = '" & JOB_NO & "' AND JOB_CHARGE_TYPE = 'M'", "")(0)

                If rowDETJOBM1.Item("MIRROR_COATING") & String.Empty = "1" Then
                    rowDETJOBM2("JOB_QTY") = IIf(rowDETJOBM1.Item("LENS_ORDER") & String.Empty = "B" Or rowDETJOBM1.Item("BALANCE_LENS") & String.Empty = "1", 2, 1)
                Else
                    rowDETJOBM2("JOB_QTY") = 0
                End If

                Dim LIST_PRICE As Decimal = Val(dst.Tables("DETPARM1").Rows(0).Item("DE_PARM_MIRROR_COATING") & String.Empty) / 2
                Dim CUST_PRICE As Decimal = LIST_PRICE

                ' See if the customer has special pricing.

                Dim sql As String = String.Empty
                sql = "SELECT * FROM DETDSGND WHERE CUST_CODE = :PARM1 AND LENS_DESIGN_CODE = :PARM2 AND SERVICE_CODE = 'M'"
                Dim rowDETDSGND As DataRow = ABSolution.ASCDATA1.GetDataRow(sql, "VV", New Object() {rowDETJOBM1.Item("CUST_CODE") & String.Empty, rowDETJOBM1.Item("LENS_DESIGN_CODE") & String.Empty})

                sql = "SELECT * FROM DETCUSTP WHERE CUST_CODE = :PARM1"
                Dim rowDETCUSTP As DataRow = ABSolution.ASCDATA1.GetDataRow(sql, "V", New Object() {rowDETJOBM1.Item("CUST_CODE") & String.Empty})

                If rowDETDSGND IsNot Nothing Then
                    CUST_PRICE = Val(rowDETDSGND.Item("CUST_PRICE") & String.Empty) / 2
                ElseIf rowDETCUSTP IsNot Nothing AndAlso rowDETCUSTP.Item("CUST_MIRROR_COATING_OVERRIDE") & String.Empty = "1" Then
                    CUST_PRICE = Val(rowDETCUSTP.Item("CUST_MIRROR_COATING_PRICE") & String.Empty) / 2
                End If

                'See if there is a price override for this mirror coating
                Dim MIRROR_COATING_COLOR As String = rowDETJOBM1.Item("MIRROR_COATING_COLOR") & String.Empty
                If MIRROR_COATING_COLOR.Length > 0 Then
                    If dst.Tables("DETCOLMC").Select("MIRROR_COATING_COLOR = '" & MIRROR_COATING_COLOR & "'").Length > 0 Then
                        Dim rowDETCOLMC As DataRow = dst.Tables("DETCOLMC").Select("MIRROR_COATING_COLOR = '" & MIRROR_COATING_COLOR & "'")(0)
                        If rowDETCOLMC IsNot Nothing AndAlso Val(rowDETCOLMC.Item("MIRROR_PRICE_OVERRIDE") & String.Empty) > 0 Then
                            CUST_PRICE = Val(rowDETCOLMC.Item("MIRROR_PRICE_OVERRIDE") & String.Empty) / 2
                        End If
                    End If
                End If

                rowDETJOBM2("LIST_PRICE") = LIST_PRICE

                Dim INV_DISC_PCT As Decimal = 0
                Dim JOB_PRICE As Decimal = CUST_PRICE * (100 - INV_DISC_PCT) / 100
                rowDETJOBM2("JOB_PRICE") = JOB_PRICE
                rowDETJOBM2("CUST_PRICE") = CUST_PRICE

            Catch ex As Exception
                RecordLogEntry("Charge_Mirror_Coating: (" & JOB_NO & ") - " & ex.Message)
            End Try

        End Sub

        Private Sub Charge_Polishing(ByVal JOB_NO As String)

            Try

                If dst.Tables("DETJOBM1").Select("JOB_NO = '" & JOB_NO & "'").Length = 0 Then
                    Exit Sub
                End If

                If dst.Tables("DETJOBM2").Select("JOB_NO = '" & JOB_NO & "'").Length = 0 Then
                    Exit Sub
                End If

                If dst.Tables("DETJOBM2").Select("JOB_NO = '" & JOB_NO & "' AND JOB_CHARGE_TYPE = 'P'", "").Length = 0 Then
                    Exit Sub
                End If

                Dim rowDETJOBM1 As DataRow = dst.Tables("DETJOBM1").Select("JOB_NO = '" & JOB_NO & "'")(0)
                Dim rowDETJOBM2 As DataRow = dst.Tables("DETJOBM2").Select("JOB_NO = '" & JOB_NO & "' AND JOB_CHARGE_TYPE = 'P'", "")(0)


                If rowDETJOBM1.Item("POLISHING") & String.Empty <> "0" Then
                    rowDETJOBM2("JOB_QTY") = IIf(rowDETJOBM1.Item("LENS_ORDER") & String.Empty = "B" Or rowDETJOBM1.Item("BALANCE_LENS") & String.Empty = "1", 2, 1)
                Else
                    rowDETJOBM2("JOB_QTY") = 0
                End If


                Dim LIST_PRICE As Decimal = 0
                Dim CUST_PRICE As Decimal = 0

                ' Validate that Polishing is a valid charge
                Dim sql As String = "Select MATL_POLISHING from DETMATL1 where MATL_CODE = :PARM1"
                Dim rowValue As DataRow = ABSolution.ASCDATA1.GetDataRow(sql, "V", New Object() {rowDETJOBM1.Item("MATL_CODE") & String.Empty})

                If rowValue IsNot Nothing Then
                    LIST_PRICE = Val(rowValue.Item("MATL_POLISHING") & String.Empty) / 2
                    CUST_PRICE = LIST_PRICE

                    ' See if the customer has special pricing.
                    sql = "SELECT * FROM DETDSGND WHERE CUST_CODE = :PARM1 AND LENS_DESIGN_CODE = :PARM2 AND SERVICE_CODE = 'P'"
                    Dim rowDETDSGND As DataRow = ABSolution.ASCDATA1.GetDataRow(sql, "VV", New Object() {rowDETJOBM1.Item("CUST_CODE") & String.Empty, rowDETJOBM1.Item("LENS_DESIGN_CODE") & String.Empty})

                    sql = "SELECT * FROM DETCUSTP WHERE CUST_CODE = :PARM1"
                    Dim rowDETCUSTP As DataRow = ABSolution.ASCDATA1.GetDataRow(sql, "V", New Object() {rowDETJOBM1.Item("CUST_CODE") & String.Empty})

                    If rowDETDSGND IsNot Nothing Then
                        CUST_PRICE = Val(rowDETDSGND.Item("CUST_PRICE") & String.Empty) / 2
                    ElseIf rowDETCUSTP IsNot Nothing AndAlso rowDETCUSTP.Item("CUST_POLISHING_OVERRIDE") & String.Empty = "1" Then
                        CUST_PRICE = Val(rowDETCUSTP.Item("CUST_POLISHING_PRICE") & String.Empty) / 2
                    End If
                End If

                rowDETJOBM2("LIST_PRICE") = LIST_PRICE

                Dim INV_DISC_PCT As Decimal = 0
                Dim JOB_PRICE As Decimal = CUST_PRICE * (100 - INV_DISC_PCT) / 100
                rowDETJOBM2("JOB_PRICE") = JOB_PRICE
                rowDETJOBM2("CUST_PRICE") = CUST_PRICE

            Catch ex As Exception
                RecordLogEntry("Charge_Polishing: (" & JOB_NO & ") - " & ex.Message)
            End Try


        End Sub

        Private Sub Charge_Tinting(ByVal JOB_NO As String)

            Try

                If dst.Tables("DETJOBM1").Select("JOB_NO = '" & JOB_NO & "'").Length = 0 Then
                    Exit Sub
                End If

                If dst.Tables("DETJOBM2").Select("JOB_NO = '" & JOB_NO & "'").Length = 0 Then
                    Exit Sub
                End If

                If dst.Tables("DETJOBM2").Select("JOB_NO = '" & JOB_NO & "' AND JOB_CHARGE_TYPE = 'T'", "").Length = 0 Then
                    Exit Sub
                End If

                Dim rowDETJOBM1 As DataRow = dst.Tables("DETJOBM1").Select("JOB_NO = '" & JOB_NO & "'")(0)
                Dim rowDETJOBM2 As DataRow = dst.Tables("DETJOBM2").Select("JOB_NO = '" & JOB_NO & "' AND JOB_CHARGE_TYPE = 'T'", "")(0)


                If rowDETJOBM1.Item("TINT_CODE") & String.Empty <> "NONE" Then
                    rowDETJOBM2("JOB_QTY") = IIf(rowDETJOBM1.Item("LENS_ORDER") & String.Empty = "B" Or rowDETJOBM1.Item("BALANCE_LENS") & String.Empty = "1", 2, 1)
                Else
                    rowDETJOBM2("JOB_QTY") = 0
                End If

                Dim LIST_PRICE As Decimal = 0
                Dim CUST_PRICE As Decimal = 0

                ' Validate that Tinting is a valid charge
                Dim sql As String = "Select TINT_PRICE from DETTINT1 where TINT_CODE = :PARM1"
                Dim rowValue As DataRow = ABSolution.ASCDATA1.GetDataRow(sql, "V", New Object() {rowDETJOBM1.Item("TINT_CODE") & String.Empty})

                If rowValue IsNot Nothing Then

                    LIST_PRICE = Val(rowValue.Item("TINT_PRICE") & String.Empty) / 2
                    CUST_PRICE = LIST_PRICE

                    ' See if the customer has special pricing.
                    sql = "SELECT * FROM DETDSGND WHERE CUST_CODE = :PARM1 AND LENS_DESIGN_CODE = :PARM2 AND SERVICE_CODE = 'T'"
                    Dim rowDETDSGND As DataRow = ABSolution.ASCDATA1.GetDataRow(sql, "VV", New Object() {rowDETJOBM1.Item("CUST_CODE") & String.Empty, rowDETJOBM1.Item("LENS_DESIGN_CODE") & String.Empty})

                    sql = "SELECT * FROM DETCUSTP WHERE CUST_CODE = :PARM1"
                    Dim rowDETCUSTP As DataRow = ABSolution.ASCDATA1.GetDataRow(sql, "V", New Object() {rowDETJOBM1.Item("CUST_CODE") & String.Empty})

                    If rowDETDSGND IsNot Nothing Then
                        CUST_PRICE = Val(rowDETDSGND.Item("CUST_PRICE") & String.Empty) / 2
                    ElseIf rowDETCUSTP IsNot Nothing AndAlso rowDETCUSTP.Item("CUST_TINT_OVERRIDE") & String.Empty = "1" Then
                        CUST_PRICE = Val(rowDETCUSTP.Item("CUST_TINT_PRICE") & String.Empty) / 2
                    End If
                End If

                rowDETJOBM2("LIST_PRICE") = LIST_PRICE

                Dim INV_DISC_PCT As Decimal = 0
                Dim JOB_PRICE As Decimal = CUST_PRICE * (100 - INV_DISC_PCT) / 100
                rowDETJOBM2("JOB_PRICE") = JOB_PRICE
                rowDETJOBM2("CUST_PRICE") = CUST_PRICE

            Catch ex As Exception
                RecordLogEntry("Charge_Tinting: (" & JOB_NO & ") - " & ex.Message)
            End Try

        End Sub

        Private Sub Charge_WrapEdge(ByVal JOB_NO As String)

            Try

                If dst.Tables("DETJOBM1").Select("JOB_NO = '" & JOB_NO & "'").Length = 0 Then
                    Exit Sub
                End If

                If dst.Tables("DETJOBM2").Select("JOB_NO = '" & JOB_NO & "'").Length = 0 Then
                    Exit Sub
                End If

                If dst.Tables("DETJOBM2").Select("JOB_NO = '" & JOB_NO & "' AND JOB_CHARGE_TYPE = 'W'", "").Length = 0 Then
                    Exit Sub
                End If

                Dim rowDETJOBM1 As DataRow = dst.Tables("DETJOBM1").Select("JOB_NO = '" & JOB_NO & "'")(0)
                Dim rowDETJOBM2 As DataRow = dst.Tables("DETJOBM2").Select("JOB_NO = '" & JOB_NO & "' AND JOB_CHARGE_TYPE = 'W'", "")(0)

                If rowDETJOBM1.Item("FINISHED") & String.Empty <> "U" And rowDETJOBM1.Item("WRAP_EDGE") & String.Empty = "1" Then
                    rowDETJOBM2("JOB_QTY") = IIf(rowDETJOBM1.Item("LENS_ORDER") & String.Empty = "B" Or rowDETJOBM1.Item("BALANCE_LENS") & String.Empty = "1", 2, 1)
                Else
                    rowDETJOBM2("JOB_QTY") = 0
                End If

                Dim LIST_PRICE As Decimal = Val(dst.Tables("DETPARM1").Rows(0).Item("DE_PARM_WRAP_EDGE") & "") / 2
                Dim CUST_PRICE As Decimal = LIST_PRICE

                ' See if the customer has special pricing.
                Dim sql As String = String.Empty
                sql = "SELECT * FROM DETDSGND WHERE CUST_CODE = :PARM1 AND LENS_DESIGN_CODE = :PARM2 AND SERVICE_CODE = 'W'"
                Dim rowDETDSGND As DataRow = ABSolution.ASCDATA1.GetDataRow(sql, "VV", New Object() {rowDETJOBM1.Item("CUST_CODE") & String.Empty, rowDETJOBM1.Item("LENS_DESIGN_CODE") & String.Empty})

                sql = "SELECT * FROM DETCUSTP WHERE CUST_CODE = :PARM1"
                Dim rowDETCUSTP As DataRow = ABSolution.ASCDATA1.GetDataRow(sql, "V", New Object() {rowDETJOBM1.Item("CUST_CODE") & String.Empty})

                If rowDETDSGND IsNot Nothing Then
                    CUST_PRICE = Val(rowDETDSGND.Item("CUST_PRICE") & String.Empty) / 2
                ElseIf rowDETCUSTP IsNot Nothing AndAlso rowDETCUSTP.Item("CUST_WRAP_EDGE_OVERRIDE") & String.Empty = "1" Then
                    CUST_PRICE = Val(rowDETCUSTP.Item("CUST_WRAP_EDGE_PRICE") & String.Empty) / 2
                End If

                rowDETJOBM2("LIST_PRICE") = LIST_PRICE

                Dim INV_DISC_PCT As Decimal = 0
                Dim JOB_PRICE As Decimal = CUST_PRICE * (100 - INV_DISC_PCT) / 100
                rowDETJOBM2("JOB_PRICE") = JOB_PRICE
                rowDETJOBM2("CUST_PRICE") = CUST_PRICE

            Catch ex As Exception
                RecordLogEntry("Charge_WrapEdge: (" & JOB_NO & ") - " & ex.Message)
            End Try

        End Sub

        Private Function BlankSelections(ByVal JOB_NO As String, ByVal clear_special_fields_if_necessary As Boolean) As Boolean

            Dim SPHERE_R As Double = 0
            Dim BASE_CURVE_R As Double = 0
            Dim OPC_CODE_R As String = ""

            Dim SPHERE_L As Double = 0
            Dim BASE_CURVE_L As Double = 0
            Dim OPC_CODE_L As String = ""

            Dim sql As String = String.Empty
            BlankSelections = True

            Try
                Dim rowDETJOBM1 As DataRow = dst.Tables("DETJOBM1").Select("JOB_NO = '" & JOB_NO & "'")(0)

                For Each rowDETJOBM3 As DataRow In dst.Tables("DETJOBM3").Select("JOB_NO = '" & JOB_NO & "'")
                    BlankSelection(JOB_NO)

                    Dim RL As String = rowDETJOBM3.Item("RL") & String.Empty

                    If RL = "R" Then
                        SPHERE_R = Val(rowDETJOBM3.Item("SPHERE") & String.Empty)
                        BASE_CURVE_R = Val(rowDETJOBM3.Item("BASE_CURVE") & String.Empty)
                        OPC_CODE_R = rowDETJOBM3.Item("OPC_CODE") & String.Empty
                    Else
                        SPHERE_L = Val(rowDETJOBM3.Item("SPHERE") & String.Empty)
                        BASE_CURVE_L = Val(rowDETJOBM3.Item("BASE_CURVE") & String.Empty)
                        OPC_CODE_L = rowDETJOBM3.Item("OPC_CODE") & String.Empty
                    End If

                    If clear_special_fields_if_necessary Then
                        If rowDETJOBM1.Item("ENTER_AS_WORN") & String.Empty <> "1" Then ' If Not Absx1.chkFor("WRAP_DESIGN").Checked Then
                            For Each COLUMN_NAME As String In New String() _
                            {"FITTING_VERTEX", "REFRACTIVE_VERTEX", "PANTOSCOPIC_TILT", "PANORAMIC_ANGLE"}
                                rowDETJOBM3.Item(COLUMN_NAME) = DBNull.Value
                            Next
                        End If

                        If rowDETJOBM1.Item("RX_PRISM") & String.Empty <> "1" Then
                            For Each COLUMN_NAME As String In New String() _
                                {"PRISM_IN", "PRISM_IN_AXIS", "PRISM_UP", "PRISM_UP_AXIS"}
                                rowDETJOBM3.Item(COLUMN_NAME) = DBNull.Value
                            Next
                        End If
                    End If
                Next

                ' this next routine should probably be used for all designers who supply their own blanks

                prevent_blank_selection = True

                Dim BASE_CURVE As Double = 0
                Dim OPC_CODE As String = ""

                If rowDETJOBM1.Item("LENS_ORDER") & String.Empty = "B" AndAlso rowDETJOBM1("LENS_DESIGNER_CODE") & String.Empty = "SEIKO" Then
                    If System.Math.Abs(SPHERE_R - SPHERE_L) <= 1 And _
                       System.Math.Round(System.Math.Abs(BASE_CURVE_R - BASE_CURVE_L), 2) <> 0 And _
                       OPC_CODE_L <> "" And OPC_CODE_R <> "" Then

                        Dim LensToFix As String
                        sql = String.Empty
                        If SPHERE_R >= 0 And SPHERE_L >= 0 Then ' use the higher Base Curve for both
                            If BASE_CURVE_R > BASE_CURVE_L Then
                                LensToFix = "L"
                                BASE_CURVE = BASE_CURVE_R
                                sql = "OPC_CODE = '" & OPC_CODE_R & "'"
                            Else
                                LensToFix = "R"
                                BASE_CURVE = BASE_CURVE_L
                                sql = "OPC_CODE_LEFT = '" & OPC_CODE_L & "'"
                            End If
                        Else ' use the lower Base Curve for both
                            If BASE_CURVE_R < BASE_CURVE_L Then
                                LensToFix = "L"
                                BASE_CURVE = BASE_CURVE_R
                                sql = "OPC_CODE = '" & OPC_CODE_R & "'"
                            Else
                                LensToFix = "R"
                                BASE_CURVE = BASE_CURVE_L
                                sql = "OPC_CODE_LEFT = '" & OPC_CODE_L & "'"
                            End If
                        End If

                        sql = "Select * from DETBLNK1 where " & sql
                        Dim rowDETBLNK1 As DataRow = ABSolution.ASCDATA1.GetDataRow(sql)
                        OPC_CODE = rowDETBLNK1.Item(IIf(LensToFix = "R", "OPC_CODE", "OPC_CODE_LEFT"))

                        For Each rowDETJOBM3 As DataRow In dst.Tables("DETJOBM3").Select("JOB_NO = '" & JOB_NO & "'")
                            If rowDETJOBM3.Item("RL") & String.Empty = LensToFix Then
                                rowDETJOBM3.Item("BASE_CURVE") = BASE_CURVE
                                rowDETJOBM3.Item("OPC_CODE") = OPC_CODE
                            End If
                        Next
                    End If
                End If
                prevent_blank_selection = False

            Catch ex As Exception
                BlankSelections = False
                RecordLogEntry("BlankSelections (" & JOB_NO & "): " & ex.Message)
            End Try

        End Function

        Private Function BlankSelection(ByVal JOB_NO As String) As Boolean

            Try

                BlankSelection = True

                If prevent_blank_selection Then Exit Function

                Dim rowDETJOBM3 As DataRow = dst.Tables("DETJOBM3").Select("JOB_NO = '" & JOB_NO & "'")(0)

                Dim MATL_CODE As String = rowDETJOBM3.Item("MATL_CODE") & String.Empty
                Dim COLOR_CODE As String = rowDETJOBM3.Item("COLOR_CODE") & String.Empty
                Dim COLOR_TYPE As String = rowDETJOBM3.Item("COLOR_TYPE") & String.Empty
                Dim LENS_DESIGN_CODE As String = rowDETJOBM3.Item("LENS_DESIGN_CODE") & String.Empty
                Dim LENS_DESIGNER_CODE As String = rowDETJOBM3.Item("LENS_DESIGNER_CODE") & String.Empty

                Dim BASE_CURVE As Double = 0
                Dim OPC_CODE As String = String.Empty

                Dim CYLINDER As Double = Val(rowDETJOBM3.Item("CYLINDER") & String.Empty)
                Dim SPHERE As Double = Val(rowDETJOBM3.Item("SPHERE") & String.Empty)
                Dim ADD_POWER As Double = Val(rowDETJOBM3.Item("ADD_POWER") & String.Empty)
                Dim RL As String = rowDETJOBM3.Item("RL") & String.Empty

                Dim SPHERE_SIGN As Decimal
                Dim SPHERE_ORIG As Double = SPHERE

                If SPHERE < 0 Then
                    SPHERE_SIGN = -1
                Else
                    SPHERE_SIGN = 1
                End If
                Dim TMP As Integer = (SPHERE / 0.25) + (SPHERE_SIGN * 0.5)

                SPHERE = TMP * 0.25

                Dim dtBlankInfo As DataTable = ABSolution.ASCDATA1.GetDataTableFromSP("DEL.BlankSelection", _
                                                                           New String() {"MATL_CODE_IN", "COLOR_CODE_IN", "COLOR_TYPE_IN", _
                                                                                         "LENS_DESIGN_CODE_IN", "LENS_DESIGNER_CODE_IN", _
                                                                                         "CYLINDER_IN", "SPHERE_IN", "ADD_POWER_IN", "RL"}, _
                                                                           New Object() {MATL_CODE, COLOR_CODE, COLOR_TYPE, LENS_DESIGN_CODE, _
                                                                                         LENS_DESIGNER_CODE, CYLINDER, SPHERE, ADD_POWER, RL})

                If dtBlankInfo.Rows.Count > 0 Then
                    rowDETJOBM3.Item("BASE_CURVE") = dtBlankInfo.Rows(0).Item("BASE_CURVE")
                    rowDETJOBM3.Item("OPC_CODE") = dtBlankInfo.Rows(0).Item("OPC_CODE")
                End If

                SPHERE = SPHERE_ORIG
                Return True

            Catch ex As Exception
                RecordLogEntry("BlankSelection (" & JOB_NO & ") " & ex.Message)
                Return False
            End Try

        End Function

#End Region

#Region "Sales Order Processing"

        ''' <summary>
        ''' Appends Chars of charToAdd to addToString if the char is not in addToString
        ''' </summary>
        ''' <param name="charsToAdd"></param>
        ''' <param name="addToString"></param>
        ''' <remarks></remarks>
        Private Sub AddCharNoDups(ByVal charsToAdd As String, ByRef addToString As String)

            Try
                If testMode Then RecordLogEntry("Enter AddCharNoDups.")

                If addToString Is Nothing Then addToString = String.Empty
                If charsToAdd Is Nothing Then charsToAdd = String.Empty

                For Each charErrorCode As Char In charsToAdd
                    If Not addToString.Contains(charErrorCode) Then
                        addToString = String.Concat(addToString, charErrorCode)
                    End If
                Next

                ' Sort the Codes
                Dim stringArray(addToString.Length) As String
                Dim i As Integer = 1

                For Each ch As Char In addToString
                    stringArray(i) = ch
                    i += 1
                Next

                Array.Sort(stringArray)

                addToString = String.Empty
                For Each ss As String In stringArray
                    addToString &= ss
                Next

                If testMode Then RecordLogEntry("Exit AddCharNoDups.")

            Catch ex As Exception
                RecordLogEntry("AddCharNoDups: " & ex.Message)
            End Try

        End Sub

        ''' <summary>
        ''' Add Sample Items to all orders where a order requests 2 or more
        ''' of an item and the item has an associated sample.
        ''' </summary>
        ''' <remarks></remarks>
        Private Sub AddSampleItemsToOrders(ByRef rowSOTORDR1 As DataRow, ByRef tblSOTORDR2 As DataTable)

            Dim ORDR_NO As String = String.Empty
            Dim ORDR_LNO As Integer = 0
            Dim rowICTITEM1 As DataRow = Nothing
            Dim errorCodes As String = String.Empty

            ORDR_NO = rowSOTORDR1.Item("ORDR_NO")

            ' This customer only
            If rowSOTORDR1.Item("CUST_CODE") <> "016917" Then
                Exit Sub
            End If

            ' Cannot be a DPD
            If rowSOTORDR1.Item("ORDR_DPD") & String.Empty = "1" Then
                Exit Sub
            End If

            ' No hold codes except Review Instructions
            If (rowSOTORDR1.Item("ORDR_REL_HOLD_CODES") & String.Empty).ToString.Trim.Replace(ReviewSpecInstr, "").Length > 0 Then
                Exit Sub
            End If

            If (rowSOTORDR1.Item("ORDR_STATUS_WEB") & String.Empty).ToString.Trim.Length > 0 _
                AndAlso (rowSOTORDR1.Item("ORDR_REL_HOLD_CODES") & String.Empty).ToString.Trim.Replace(ReviewSpecInstr, "").Length > 0 Then
                Exit Sub
            End If

            ORDR_LNO = Val(dst.Tables("SOTORDR2").Compute("MAX(ORDR_LNO)", "ORDR_NO = '" & ORDR_NO & "'")) + 1

            For Each rowSOTORDR2 As DataRow In tblSOTORDR2.Select("ORDR_NO = '" & ORDR_NO & "' AND ORDR_QTY >= 2 AND ISNLL(SAMPLE_IND, '0') = '0'")

                rowICTITEM1 = GetSampleItem(rowSOTORDR2.Item("ITEM_CODE") & String.Empty)
                If rowICTITEM1 Is Nothing Then Continue For
                If rowICTITEM1.Item("ITEM_CODE") & String.Empty = String.Empty Then Continue For

                rowSOTORDR2 = dst.Tables("SOTORDR2").NewRow
                rowSOTORDR2.Item("ORDR_NO") = ORDR_NO
                rowSOTORDR2.Item("ORDR_LNO") = ORDR_LNO
                ORDR_LNO += 1
                rowSOTORDR2.Item("ITEM_CODE") = rowICTITEM1.Item("ITEM_CODE") & String.Empty
                rowSOTORDR2.Item("ITEM_DESC") = rowICTITEM1.Item("ITEM_DESC") & String.Empty
                rowSOTORDR2.Item("ITEM_DESC2") = rowICTITEM1.Item("ITEM_DESC2") & String.Empty
                dst.Tables("SOTORDR2").Rows.Add(rowSOTORDR2)

                SetItemInfo(rowSOTORDR2, errorCodes)

                rowSOTORDR2.Item("ORDR_UNIT_PRICE") = 0
                rowSOTORDR2.Item("ORDR_QTY") = 1
                rowSOTORDR2.Item("ORDR_QTY_OPEN") = 1
                rowSOTORDR2.Item("ORDR_QTY_ORIG") = 1
                rowSOTORDR2.Item("ORDR_LINE_STATUS") = "O"
                rowSOTORDR2.Item("CUST_LINE_REF") = rowSOTORDR2.Item("CUST_LINE_REF") & String.Empty
                rowSOTORDR2.Item("ORDR_LINE_SOURCE") = rowSOTORDR2.Item("ORDR_LINE_SOURCE") & String.Empty
                rowSOTORDR2.Item("ORDR_LR") = rowSOTORDR2.Item("ORDR_LR") & String.Empty

                rowSOTORDR2.Item("ORDR_QTY_PICK") = 0
                rowSOTORDR2.Item("ORDR_QTY_SHIP") = 0
                rowSOTORDR2.Item("ORDR_QTY_CANC") = 0
                rowSOTORDR2.Item("ORDR_QTY_BACK") = 0
                rowSOTORDR2.Item("ORDR_QTY_ONPO") = 0

                rowSOTORDR2.Item("PATIENT_NAME") = rowSOTORDR2.Item("PATIENT_NAME") & String.Empty
                rowSOTORDR2.Item("ORDR_UNIT_PRICE_OVERRIDDEN") = String.Empty
                rowSOTORDR2.Item("ORDR_QTY_DFCT") = 0
                rowSOTORDR2.Item("ORDR_UNIT_PRICE_PATIENT") = 0
                rowSOTORDR2.Item("HANDLING_CODE") = String.Empty
                rowSOTORDR2.Item("SAMPLE_IND") = "1"
                rowSOTORDR2.Item("ORDR_REL_HOLD_CODES") = String.Empty

                dst.Tables("SOTORDR2").Rows.Add(rowSOTORDR2)
            Next

        End Sub

        Private Function CreateCustomerShipTo(ByVal CUST_CODE As String, ByVal CUST_SHIP_TO_NO As String _
                                            , ByRef rowSOTORDRX As DataRow, ByRef rowARTCUST1 As DataRow) As Boolean

            Try
                If testMode Then RecordLogEntry("Enter CreateCustomerShipTo: " & CUST_CODE)

                If rowARTCUST1 Is Nothing OrElse (rowARTCUST1.Item("CUST_CODE") & String.Empty).ToString.Trim.Length = 0 Then
                    Return True
                End If

                Dim sql As String = String.Empty
                Dim CUST_SHIP_TO_SHIP_VIA_CODE As String = String.Empty
                Dim CUST_SHIP_TO_STATE As String = (rowSOTORDRX.Item("CUST_SHIP_TO_STATE") & String.Empty).ToString.Replace("'", "").Trim
                Dim STAX_EXEMPT As String = (rowARTCUST1.Item("STAX_EXEMPT") & String.Empty).ToString.Trim
                If STAX_EXEMPT.Length = 0 Then STAX_EXEMPT = "0"

                ' See if we have predetermined Ship vias for the customer / state combination
                Dim rowSOTSVIAE As DataRow = ABSolution.ASCDATA1.GetDataRow("SELECT * FROM SOTSVIAE WHERE CUST_CODE = '" & CUST_CODE & "' AND STATE_CODE = '" & CUST_SHIP_TO_STATE & "'")
                If rowSOTSVIAE IsNot Nothing Then
                    CUST_SHIP_TO_SHIP_VIA_CODE = (rowSOTSVIAE.Item("SHIP_VIA_CODE") & String.Empty).ToString.Trim
                End If

                Dim CUST_SHIP_TO_PHONE As String = String.Empty
                For Each chPhone As Char In rowSOTORDRX.Item("CUST_SHIP_TO_PHONE") & String.Empty
                    If Char.IsDigit(chPhone) Then
                        CUST_SHIP_TO_PHONE &= chPhone
                    End If
                Next
                CUST_SHIP_TO_PHONE = TruncateField(CUST_SHIP_TO_PHONE, "ARTCUST2", "CUST_SHIP_TO_PHONE")

                Dim CUST_SHIP_TO_URL As String = TruncateField(rowSOTORDRX.Item("OFFICE_WEBSITE") & String.Empty, "ARTCUST2", "CUST_SHIP_TO_URL").ToLower

                sql = "INSERT INTO ARTCUST2 "
                sql &= " ("
                sql &= "CUST_CODE, CUST_SHIP_TO_NO, CUST_SHIP_TO_NAME,"
                sql &= " CUST_SHIP_TO_ADDR1, CUST_SHIP_TO_ADDR2, CUST_SHIP_TO_ADDR3,"
                sql &= " CUST_SHIP_TO_CITY, CUST_SHIP_TO_STATE, CUST_SHIP_TO_ZIP_CODE,"
                sql &= " CUST_SHIP_TO_COUNTRY, INIT_DATE, INIT_OPER,"
                sql &= " LAST_DATE, LAST_OPER, CUST_SHIP_TO_STATUS, CUST_SHIP_TO_SHIP_VIA_CODE, STAX_EXEMPT,"
                sql &= " CUST_SHIP_TO_PHONE, CUST_SHIP_TO_URL"
                sql &= ")"
                sql &= " VALUES "
                sql &= " ("
                sql &= "'" & CUST_CODE & "'"
                sql &= ", '" & CUST_SHIP_TO_NO & "'"
                sql &= ", '" & (rowSOTORDRX.Item("CUST_SHIP_TO_NAME") & String.Empty).ToString.Replace("'", "").Trim & "'"
                sql &= ", '" & (rowSOTORDRX.Item("CUST_SHIP_TO_ADDR1") & String.Empty).ToString.Replace("'", "").Trim & "'"
                sql &= ", '" & (rowSOTORDRX.Item("CUST_SHIP_TO_ADDR2") & String.Empty).ToString.Replace("'", "").Trim & "'"
                sql &= ", '" & (rowSOTORDRX.Item("CUST_SHIP_TO_ADDR3") & String.Empty).ToString.Replace("'", "").Trim & "'"
                sql &= ", '" & (rowSOTORDRX.Item("CUST_SHIP_TO_CITY") & String.Empty).ToString.Replace("'", "").Trim & "'"
                sql &= ", '" & (rowSOTORDRX.Item("CUST_SHIP_TO_STATE") & String.Empty).ToString.Replace("'", "").Trim & "'"
                sql &= ", '" & (rowSOTORDRX.Item("CUST_SHIP_TO_ZIP_CODE") & String.Empty).ToString.Replace("'", "").Trim & "'"
                sql &= ", '" & rowSOTORDRX.Item("CUST_SHIP_TO_COUNTRY") & String.Empty & "'"
                sql &= ", SYSDATE"
                sql &= ", '" & ABSolution.ASCMAIN1.USER_ID & "'"
                sql &= ", SYSDATE"
                sql &= ", '" & ABSolution.ASCMAIN1.USER_ID & "'"
                sql &= ", 'A'"
                sql &= ", '" & CUST_SHIP_TO_SHIP_VIA_CODE & "'"
                sql &= ", '" & STAX_EXEMPT & "'"
                sql &= ", '" & CUST_SHIP_TO_PHONE & "'"
                sql &= ", '" & CUST_SHIP_TO_URL & "'"
                sql &= ")"

                ABSolution.ASCDATA1.ExecuteSQL(sql)
                System.Threading.Thread.Sleep(2000)
                rowARTCUST2 = baseClass.LookUp("ARTCUST2", New String() {CUST_CODE, CUST_SHIP_TO_NO})

                If testMode Then RecordLogEntry("Exit CreateCustomerShipTo: " & CUST_CODE)

                Return True
            Catch ex As Exception
                RecordLogEntry("CreateCustomerShipTo: (" & CUST_CODE & ") " & ex.Message)
                Return False
            End Try

        End Function

        Private Function CreateOrderBillTo(ByVal ORDR_NO As String) As Boolean

            Dim rowSOTORDR5 As DataRow = Nothing
            Try

                If testMode Then RecordLogEntry("Exter CreateOrderBillTo: " & ORDR_NO)

                rowSOTORDR5 = dst.Tables("SOTORDR5").NewRow

                rowSOTORDR5.Item("ORDR_NO") = ORDR_NO
                rowSOTORDR5.Item("CUST_ADDR_TYPE") = "BT"
                dst.Tables("SOTORDR5").Rows.Add(rowSOTORDR5)

                ' Bill To
                If rowARTCUST1 IsNot Nothing Then
                    rowSOTORDR5.Item("CUST_NAME") = rowARTCUST1.Item("CUST_NAME") & String.Empty
                    rowSOTORDR5.Item("CUST_ADDR1") = rowARTCUST1.Item("CUST_ADDR1") & String.Empty
                    rowSOTORDR5.Item("CUST_ADDR2") = rowARTCUST1.Item("CUST_ADDR2") & String.Empty
                    rowSOTORDR5.Item("CUST_ADDR3") = rowARTCUST1.Item("CUST_ADDR3") & String.Empty
                    rowSOTORDR5.Item("CUST_CITY") = rowARTCUST1.Item("CUST_CITY") & String.Empty
                    rowSOTORDR5.Item("CUST_STATE") = rowARTCUST1.Item("CUST_STATE") & String.Empty
                    rowSOTORDR5.Item("CUST_ZIP_CODE") = rowARTCUST1.Item("CUST_ZIP_CODE") & String.Empty
                    rowSOTORDR5.Item("CUST_COUNTRY") = rowARTCUST1.Item("CUST_COUNTRY") & String.Empty
                    rowSOTORDR5.Item("CUST_CONTACT") = rowARTCUST1.Item("CUST_CONTACT") & String.Empty
                    rowSOTORDR5.Item("CUST_PHONE") = rowARTCUST1.Item("CUST_PHONE") & String.Empty
                    rowSOTORDR5.Item("CUST_EXT") = rowARTCUST1.Item("CUST_EXT") & String.Empty
                    rowSOTORDR5.Item("CUST_FAX") = rowARTCUST1.Item("CUST_FAX") & String.Empty
                    rowSOTORDR5.Item("CUST_EMAIL") = rowARTCUST1.Item("CUST_EMAIL") & String.Empty
                End If

                If testMode Then RecordLogEntry("Exit CreateOrderBillTo: " & ORDR_NO)
                Return True

            Catch ex As Exception
                RecordLogEntry("CreateOrderBillTo: (" & ORDR_NO & ") " & ex.Message)
                Return False
            End Try

        End Function

        ''' <summary>
        ''' Creates and Entry in SOTORDRW for errors on the order
        ''' </summary>
        ''' <param name="ErrorCode"></param>
        ''' <param name="ORDR_LNO"></param>
        ''' <remarks></remarks>
        Private Sub CreateOrderErrorRecord(ByVal ORDR_NO As String, ByVal ORDR_LNO As Integer, ByVal ErrorCode As String, Optional ByVal ErrorMessage As String = "")

            Try
                If testMode Then RecordLogEntry("Enter CreateOrderErrorRecord: " & ORDR_NO)

                If dst.Tables("SOTORDRW").Select("ORDR_NO = '" & ORDR_NO & "' AND ORDR_LNO = " & ORDR_LNO & " AND ORDR_ERROR_CODE = '" & ErrorCode & "'").Length > 0 Then
                    dst.Tables("SOTORDRW").Select("ORDR_NO = '" & ORDR_NO & "' AND ORDR_LNO = " & ORDR_LNO & " AND ORDR_ERROR_CODE = '" & ErrorCode & "'")(0).Delete()
                End If

                Dim rowSOTORDRW As DataRow = dst.Tables("SOTORDRW").NewRow

                rowSOTORDRW.Item("ORDR_NO") = ORDR_NO
                rowSOTORDRW.Item("ORDR_LNO") = ORDR_LNO
                rowSOTORDRW.Item("ORDR_ERROR_CODE") = ErrorCode

                If ErrorMessage.Length > 0 Then
                    rowSOTORDRW.Item("ORDR_ERROR_TEXT") = ErrorMessage
                ElseIf dst.Tables("SOTORDRO").Select("ORDR_REL_HOLD_CODES = '" & ErrorCode & "'").Length > 0 Then
                    rowSOTORDRW.Item("ORDR_ERROR_TEXT") = dst.Tables("SOTORDRO").Select("ORDR_REL_HOLD_CODES = '" & ErrorCode & "'")(0).Item("ORDR_COMMENT")
                Else
                    rowSOTORDRW.Item("ORDR_ERROR_TEXT") = "Unknown Error"
                End If

                dst.Tables("SOTORDRW").Rows.Add(rowSOTORDRW)
                If testMode Then RecordLogEntry("Exit CreateOrderErrorRecord: " & ORDR_NO)

            Catch ex As Exception
                RecordLogEntry("CreateOrderErrorRecord: " & ex.Message)
            End Try

        End Sub

        Private Function CreateOrderShipTo(ByVal ORDR_NO As String, ByRef rowSOTORDRX As DataRow, ByVal UseImportedShipToAddress As Boolean) As Boolean

            Dim rowSOTORDR5 As DataRow = Nothing

            Try

                If testMode Then RecordLogEntry("Enter CreateOrderShipTo: " & ORDR_NO)

                rowSOTORDR5 = dst.Tables("SOTORDR5").NewRow
                rowSOTORDR5.Item("ORDR_NO") = ORDR_NO
                rowSOTORDR5.Item("CUST_ADDR_TYPE") = "ST"
                dst.Tables("SOTORDR5").Rows.Add(rowSOTORDR5)

                If UseImportedShipToAddress Then
                    rowSOTORDR5.Item("CUST_NAME") = rowSOTORDRX.Item("CUST_NAME") & String.Empty
                    rowSOTORDR5.Item("CUST_ADDR1") = rowSOTORDRX.Item("CUST_ADDR1") & String.Empty
                    rowSOTORDR5.Item("CUST_ADDR2") = rowSOTORDRX.Item("CUST_ADDR2") & String.Empty
                    rowSOTORDR5.Item("CUST_CITY") = rowSOTORDRX.Item("CUST_CITY") & String.Empty
                    rowSOTORDR5.Item("CUST_STATE") = rowSOTORDRX.Item("CUST_STATE") & String.Empty
                    rowSOTORDR5.Item("CUST_ZIP_CODE") = rowSOTORDRX.Item("CUST_ZIP_CODE") & String.Empty
                    rowSOTORDR5.Item("CUST_COUNTRY") = "US"
                    rowSOTORDR5.Item("CUST_CONTACT") = String.Empty

                    ' Remove trailing 0000 on a Zip Code
                    If (rowSOTORDR5.Item("CUST_ZIP_CODE") & String.Empty).ToString.Length = 9 Then
                        If (rowSOTORDR5.Item("CUST_ZIP_CODE") & String.Empty).ToString.Substring(5) = "0000" Then
                            rowSOTORDR5.Item("CUST_ZIP_CODE") = rowSOTORDR5.Item("CUST_ZIP_CODE").ToString.Substring(0, 5)
                        End If
                    End If

                    rowSOTORDR5.Item("CUST_PHONE") = rowSOTORDRX.Item("CUST_PHONE") & String.Empty
                    rowSOTORDR5.Item("CUST_EXT") = String.Empty
                    rowSOTORDR5.Item("CUST_FAX") = rowSOTORDRX.Item("CUST_FAX") & String.Empty
                    rowSOTORDR5.Item("CUST_EMAIL") = rowSOTORDRX.Item("CUST_EMAIL") & String.Empty

                ElseIf rowARTCUST2 IsNot Nothing Then
                    rowSOTORDR5.Item("CUST_NAME") = rowARTCUST2.Item("CUST_SHIP_TO_NAME") & String.Empty
                    rowSOTORDR5.Item("CUST_ADDR1") = rowARTCUST2.Item("CUST_SHIP_TO_ADDR1") & String.Empty
                    rowSOTORDR5.Item("CUST_ADDR2") = rowARTCUST2.Item("CUST_SHIP_TO_ADDR2") & String.Empty
                    rowSOTORDR5.Item("CUST_ADDR3") = rowARTCUST2.Item("CUST_SHIP_TO_ADDR3") & String.Empty
                    rowSOTORDR5.Item("CUST_CITY") = rowARTCUST2.Item("CUST_SHIP_TO_CITY") & String.Empty
                    rowSOTORDR5.Item("CUST_STATE") = rowARTCUST2.Item("CUST_SHIP_TO_STATE") & String.Empty
                    rowSOTORDR5.Item("CUST_ZIP_CODE") = rowARTCUST2.Item("CUST_SHIP_TO_ZIP_CODE") & String.Empty
                    rowSOTORDR5.Item("CUST_COUNTRY") = rowARTCUST2.Item("CUST_SHIP_TO_COUNTRY") & String.Empty
                    rowSOTORDR5.Item("CUST_CONTACT") = rowARTCUST2.Item("CUST_SHIP_TO_CONTACT") & String.Empty
                    rowSOTORDR5.Item("CUST_PHONE") = rowARTCUST2.Item("CUST_SHIP_TO_PHONE") & String.Empty
                    rowSOTORDR5.Item("CUST_EXT") = rowARTCUST2.Item("CUST_SHIP_TO_EXT") & String.Empty
                    rowSOTORDR5.Item("CUST_FAX") = rowARTCUST2.Item("CUST_SHIP_TO_FAX") & String.Empty
                    rowSOTORDR5.Item("CUST_EMAIL") = rowARTCUST2.Item("CUST_SHIP_TO_EMAIL") & String.Empty

                ElseIf rowARTCUST1 IsNot Nothing Then
                    rowSOTORDR5.Item("CUST_NAME") = rowARTCUST1.Item("CUST_NAME") & String.Empty
                    rowSOTORDR5.Item("CUST_ADDR1") = rowARTCUST1.Item("CUST_ADDR1") & String.Empty
                    rowSOTORDR5.Item("CUST_ADDR2") = rowARTCUST1.Item("CUST_ADDR2") & String.Empty
                    rowSOTORDR5.Item("CUST_ADDR3") = rowARTCUST1.Item("CUST_ADDR3") & String.Empty
                    rowSOTORDR5.Item("CUST_CITY") = rowARTCUST1.Item("CUST_CITY") & String.Empty
                    rowSOTORDR5.Item("CUST_STATE") = rowARTCUST1.Item("CUST_STATE") & String.Empty
                    rowSOTORDR5.Item("CUST_ZIP_CODE") = rowARTCUST1.Item("CUST_ZIP_CODE") & String.Empty
                    rowSOTORDR5.Item("CUST_COUNTRY") = rowARTCUST1.Item("CUST_COUNTRY") & String.Empty
                    rowSOTORDR5.Item("CUST_CONTACT") = rowARTCUST1.Item("CUST_CONTACT") & String.Empty
                    rowSOTORDR5.Item("CUST_PHONE") = rowARTCUST1.Item("CUST_PHONE") & String.Empty
                    rowSOTORDR5.Item("CUST_EXT") = rowARTCUST1.Item("CUST_EXT") & String.Empty
                    rowSOTORDR5.Item("CUST_FAX") = rowARTCUST1.Item("CUST_FAX") & String.Empty
                    rowSOTORDR5.Item("CUST_EMAIL") = rowARTCUST1.Item("CUST_EMAIL") & String.Empty
                End If

                If testMode Then RecordLogEntry("Exit CreateOrderShipTo: " & ORDR_NO)
                Return True

            Catch ex As Exception
                RecordLogEntry("CreateOrderShipTo: (" & ORDR_NO & ") " & ex.Message)
                Return False
            End Try

        End Function

        ''' <summary>
        ''' Creates Sales Order
        ''' </summary>
        ''' <param name="ORDR_NO">Order Number created by this procedure</param>
        ''' <param name="CreateShipTo">Should this procedure create the ship to if if does not exist</param>
        ''' <param name="LocateShipToByTelephone">Use the customer telephone number to located the ship to</param>
        ''' <param name="ORDR_LINE_SOURCE">Order Line Source code for detail records</param>
        ''' <param name="ORDR_SOURCE">Order Source on the Sales Order header record</param>
        ''' <param name="CALLER_NAME">Caller name on header record</param>
        ''' <param name="AlwaysUseImportedShipToAddress">Use the Ship To Address supplied by the Import</param>
        ''' <param name="ORDR_STATUS_WEB">Code to place in the field that marks the record as having an error</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Private Function CreateSalesOrder(ByRef ORDR_NO As String, ByVal CreateShipTo As Boolean, _
                                          ByVal LocateShipToByTelephone As Boolean, ByVal ORDR_LINE_SOURCE As String, _
                                          ByVal ORDR_SOURCE As String, ByVal CALLER_NAME As String, _
                                          ByVal AlwaysUseImportedShipToAddress As Boolean, _
                                          ByVal ORDR_STATUS_WEB As String) As Boolean

            Try

                If testMode Then RecordLogEntry("Enter CreateSalesOrder: " & ORDR_NO)

                CreateSalesOrder = False
                ORDR_NO = String.Empty

                ' See if we have any data to process
                If dst.Tables("SOTORDRX").Rows.Count = 0 Then
                    Return False
                End If

                'Dim ORDR_NO As String = String.Empty
                Dim CUST_CODE As String = String.Empty
                Dim CUST_SHIP_TO_NO As String = String.Empty
                Dim CUST_SHIP_TO_PHONE As String = String.Empty

                Dim ORDR_QTY As Integer = 0
                Dim shipToPatient As Boolean = False
                Dim ITEM_CODE As String = String.Empty

                Dim sql As String = String.Empty
                Dim sqlSalesOrder As String = String.Empty

                Dim SOTORDR1ErrorCodes As String = String.Empty
                Dim SOTORDR2ErrorCodes As String = String.Empty

                Dim rowSOTORDR1 As DataRow = Nothing
                Dim rowSOTORDR2 As DataRow = Nothing
                Dim rowSOTORDR3 As DataRow = Nothing
                Dim rowICTPCAT1 As DataRow = Nothing

                Dim errorCodes As String = String.Empty

                Dim fieldValue As String = String.Empty
                Dim maxLength As Int16 = 0
                Dim iRow As Integer = 2
                Const unKnownShipTo As String = "xxxxxx"

                ORDR_NO = ABSolution.ASCMAIN1.Next_Control_No("SOTORDR1.ORDR_NO", 1)
                ' Grab the inital row to create the header record
                Dim rowSOTORDRX As DataRow = dst.Tables("SOTORDRX").Select("", "ORDR_LNO")(0)

                rowARTCUST1 = Nothing
                rowARTCUST2 = Nothing
                rowARTCUST3 = Nothing

                CUST_CODE = (rowSOTORDRX.Item("CUST_CODE") & String.Empty).ToString.Trim
                CUST_SHIP_TO_NO = (rowSOTORDRX.Item("CUST_SHIP_TO_NO") & String.Empty).ToString.Trim
                CUST_SHIP_TO_PHONE = (rowSOTORDRX.Item("CUST_SHIP_TO_PHONE") & String.Empty).ToString.Trim

                CUST_CODE = ABSolution.ASCMAIN1.Format_Field(CUST_CODE, "CUST_CODE")

                If CUST_CODE.Length = 0 Then
                    CUST_CODE = (rowSOTORDRX.Item("CUST_CODE") & String.Empty).ToString.Trim
                    CUST_CODE = TruncateField(CUST_CODE, "SOTORDR1", "CUST_CODE")
                End If

                If CUST_SHIP_TO_NO.Length > 0 Then
                    CUST_SHIP_TO_NO = ABSolution.ASCMAIN1.Format_Field(CUST_SHIP_TO_NO, "CUST_SHIP_TO_NO")
                End If

                rowARTCUST1 = baseClass.LookUp("ARTCUST1", CUST_CODE)
                rowARTCUST2 = baseClass.LookUp("ARTCUST2", New String() {CUST_CODE, CUST_SHIP_TO_NO})
                rowARTCUST3 = baseClass.LookUp("ARTCUST3", CUST_CODE)

                If rowARTCUST2 Is Nothing AndAlso rowARTCUST1 IsNot Nothing AndAlso CreateShipTo = True AndAlso CUST_SHIP_TO_NO.Length > 0 Then
                    CreateCustomerShipTo(CUST_CODE, CUST_SHIP_TO_NO, rowSOTORDRX, rowARTCUST1)
                ElseIf rowARTCUST2 IsNot Nothing AndAlso rowARTCUST1 IsNot Nothing AndAlso CreateShipTo = True AndAlso CUST_SHIP_TO_NO.Length > 0 Then
                    UpdateCustomerShipTo(CUST_CODE, CUST_SHIP_TO_NO, rowSOTORDRX, rowARTCUST2)
                ElseIf rowARTCUST1 IsNot Nothing AndAlso LocateShipToByTelephone = True Then
                    sql = "SELECT * From ARTCUST2 WHERE CUST_CODE = :PARM1 AND CUST_SHIP_TO_PHONE = :PARM2"
                    rowARTCUST2 = ABSolution.ASCDATA1.GetDataRow(sql, "VV", New Object() {CUST_CODE, CUST_SHIP_TO_PHONE})
                    If rowARTCUST2 IsNot Nothing Then
                        CUST_SHIP_TO_NO = rowARTCUST2.Item("CUST_SHIP_TO_NO") & String.Empty
                    ElseIf rowARTCUST1.Item("CUST_PHONE") & String.Empty = CUST_SHIP_TO_PHONE Then
                        CUST_SHIP_TO_NO = String.Empty
                    Else
                        CUST_SHIP_TO_NO = unKnownShipTo
                    End If
                End If

                rowSOTORDR1 = dst.Tables("SOTORDR1").NewRow
                rowSOTORDR1.Item("ORDR_NO") = ORDR_NO
                rowSOTORDR1.Item("ORDR_TYPE_CODE") = rowSOTORDRX.Item("ORDR_TYPE_CODE") & String.Empty
                rowSOTORDR1.Item("CUST_CODE") = CUST_CODE
                rowSOTORDR1.Item("CUST_SHIP_TO_NO") = CUST_SHIP_TO_NO
                rowSOTORDR1.Item("ORDR_STATUS") = "O"
                rowSOTORDR1.Item("ORDR_LOCK_SHIP_VIA") = (rowSOTORDRX.Item("ORDR_LOCK_SHIP_VIA") & String.Empty).ToString.Trim
                dst.Tables("SOTORDR1").Rows.Add(rowSOTORDR1)

                If ORDR_SOURCE.Length = 0 Then
                    ORDR_SOURCE = "K"
                End If
                rowSOTORDR1.Item("ORDR_SOURCE") = ORDR_SOURCE

                If rowSOTORDRX.Item("ORDR_DPD") & String.Empty = "1" Then
                    rowSOTORDR1.Item("ORDR_DPD") = "1"
                    shipToPatient = True
                Else
                    rowSOTORDR1.Item("ORDR_DPD") = "0"
                    shipToPatient = False
                End If

                If rowARTCUST1 IsNot Nothing AndAlso Not shipToPatient Then
                    rowSOTORDR1.Item("ORDR_COD_ADDON_AMT") = Val(rowARTCUST1.Item("CUST_COD_ADDON_AMT") & String.Empty)
                End If

                If rowARTCUST1 IsNot Nothing Then
                    If rowARTCUST1.Item("CUST_SHIP_TO_NO_REQD") & String.Empty = "1" Then
                        If rowARTCUST2 Is Nothing Then
                            Me.AddCharNoDups(RequiresShipTo, SOTORDR1ErrorCodes)
                        End If
                    End If
                End If

                errorCodes = String.Empty

                ' Validate the Ship Via Code
                If (rowSOTORDRX.Item("SHIP_VIA_CODE") & String.Empty).ToString.Trim.ToUpper = "STANDARD" Then
                    rowSOTORDRX.Item("SHIP_VIA_CODE") = String.Empty
                End If

                rowSOTSVIA1 = baseClass.LookUp("SOTSVIA1", (rowSOTORDRX.Item("SHIP_VIA_CODE") & String.Empty))

                If rowSOTSVIA1 Is Nothing Then
                    Dim rowSOTSVIAF As DataRow = ABSolution.ASCDATA1.GetDataRow("SELECT * FROM SOTSVIAF WHERE ORDR_SOURCE = :PARM1 AND UPPER(SHIP_VIA_DESC) = :PARM2", _
                                                                     "VV", New String() {ORDR_LINE_SOURCE, (rowSOTORDRX.Item("SHIP_VIA_CODE") & String.Empty).ToString.ToUpper.Trim})
                    If rowSOTSVIAF IsNot Nothing Then
                        If rowSOTORDR1.Item("ORDR_DPD") = "1" Then
                            rowSOTORDR1.Item("SHIP_VIA_CODE") = rowSOTSVIAF.Item("SHIP_VIA_CODE_DPD") & String.Empty
                        Else
                            rowSOTORDR1.Item("SHIP_VIA_CODE") = rowSOTSVIAF.Item("SHIP_VIA_CODE") & String.Empty
                        End If
                    End If
                Else
                    rowSOTORDR1.Item("SHIP_VIA_CODE") = (rowSOTORDRX.Item("SHIP_VIA_CODE") & String.Empty).ToString.Trim
                End If

                ' Grab the customer settings
                rowSOTORDR1.Item("CUST_NAME") = (rowSOTORDRX.Item("CUST_NAME") & String.Empty).ToString.Trim
                If Not Me.SetBillToAttributes(CUST_CODE, CUST_SHIP_TO_NO, rowSOTORDR1, errorCodes) Then
                    Me.AddCharNoDups(InvalidSoldTo, errorCodes)
                End If

                Me.AddCharNoDups(errorCodes, SOTORDR1ErrorCodes)

                ' Ship To
                If CUST_SHIP_TO_NO.Length > 0 Then
                    If Not Me.SetShipToAttributes(rowSOTORDR1, SOTORDR1ErrorCodes) Then
                        Me.AddCharNoDups(InvalidShipTo, SOTORDR1ErrorCodes)
                    End If
                End If

                ' Set DPD settings if there is no ship Via
                If shipToPatient = True AndAlso (rowSOTORDR1.Item("SHIP_VIA_CODE") & String.Empty) = String.Empty Then
                    If Not Me.SetDPDShipViaSettings(rowSOTORDR1, SOTORDR1ErrorCodes, ORDR_LINE_SOURCE) Then
                        Me.AddCharNoDups(InvalidDPD, SOTORDR1ErrorCodes)
                    End If
                End If

                If (rowSOTORDR1.Item("SHIP_VIA_CODE") & String.Empty).ToString.Trim.Length = 0 Then
                    Me.AddCharNoDups(InvalidShipVia, SOTORDR1ErrorCodes)
                End If

                ' ************ Other Order Header Fields ************
                If IsDate(rowSOTORDRX.Item("ORDR_DATE") & String.Empty) Then
                    rowSOTORDR1.Item("ORDR_DATE") = CDate(rowSOTORDRX.Item("ORDR_DATE") & String.Empty).ToString("dd-MMM-yyyy")
                Else
                    rowSOTORDR1.Item("ORDR_DATE") = DateTime.Now.ToString("dd-MMM-yyyy")
                End If

                rowSOTORDR1.Item("ORDR_CUST_PO") = rowSOTORDRX.Item("ORDR_CUST_PO") & String.Empty
                rowSOTORDR1.Item("BRANCH_CODE") = "NY"
                rowSOTORDR1.Item("DIVISION_CODE") = "ODG"
                rowSOTORDR1.Item("WHSE_CODE") = "001"

                rowSOTORDR1.Item("ORDR_PICK_SEQ") = 0
                rowSOTORDR1.Item("ORDR_HOLD_SALES") = String.Empty
                rowSOTORDR1.Item("ORDR_HOLD_CREDIT") = String.Empty
                rowSOTORDR1.Item("ORDR_HOLD_CREDIT_REL_BY") = String.Empty
                rowSOTORDR1.Item("ORDR_HOLD_CREDIT_REL_NOTE") = String.Empty
                rowSOTORDR1.Item("ORDR_CLOSED_BY") = String.Empty
                rowSOTORDR1.Item("ORDR_REV_NO") = 1
                rowSOTORDR1.Item("ORDR_HOLD_CREDIT_REASON") = String.Empty
                rowSOTORDR1.Item("ORDR_NO_FREIGHT") = "0"
                rowSOTORDR1.Item("MISC_CHG_CODE") = String.Empty
                rowSOTORDR1.Item("ORDR_NO_WEB") = String.Empty
                rowSOTORDR1.Item("ORDR_MISC_CHG_AMT") = 0
                rowSOTORDR1.Item("PATIENT_NO") = String.Empty
                rowSOTORDR1.Item("ORDR_COMMENT") = String.Empty
                rowSOTORDR1.Item("EDI_CUST_REF_NO") = rowSOTORDRX.Item("EDI_CUST_REF_NO") & String.Empty
                rowSOTORDR1.Item("ORDR_CALLER_NAME") = CALLER_NAME

                Dim ORDR_COMMENT As String = (rowSOTORDRX.Item("ORDR_COMMENT") & String.Empty).ToString.Trim

                If LocateShipToByTelephone = True AndAlso CUST_SHIP_TO_NO = unKnownShipTo Then
                    ORDR_COMMENT = "Ship To Telephone: " & CUST_SHIP_TO_PHONE & ", " & ORDR_COMMENT
                End If

                If ORDR_COMMENT.Length > 0 Then
                    Me.AddCharNoDups(ReviewSpecInstr, SOTORDR1ErrorCodes)

                    rowSOTORDR1.Item("ORDR_COMMENT") = TruncateField(ORDR_COMMENT, "SOTORDR1", "ORDR_COMMENT")

                    Dim ORDR_TNO As Integer = 0
                    While ORDR_COMMENT.Length > 0
                        ORDR_TNO += 1
                        rowSOTORDR3 = dst.Tables("SOTORDR3").NewRow
                        rowSOTORDR3.Item("ORDR_NO") = ORDR_NO
                        rowSOTORDR3.Item("ORDR_LNO") = 0
                        rowSOTORDR3.Item("ORDR_TNO") = ORDR_TNO
                        rowSOTORDR3.Item("ORDR_TEXT_PICK") = "1"
                        rowSOTORDR3.Item("ORDR_TEXT_INV") = "0"
                        dst.Tables("SOTORDR3").Rows.Add(rowSOTORDR3)

                        If ORDR_COMMENT.Length <= 100 Then
                            rowSOTORDR3.Item("ORDR_TEXT") = ORDR_COMMENT
                            ORDR_COMMENT = String.Empty
                        Else
                            rowSOTORDR3.Item("ORDR_TEXT") = ORDR_COMMENT.Substring(0, 100)
                            ORDR_COMMENT = ORDR_COMMENT.Substring(100)
                        End If
                    End While
                End If

                Me.Record_Event(ORDR_NO, "Imported Order Received.")

                ' ************ ORDER DETAILS ************
                For Each rowImportDetails As DataRow In dst.Tables("SOTORDRX").Select(sqlSalesOrder, "ORDR_LNO")

                    SOTORDR2ErrorCodes = String.Empty

                    rowSOTORDR2 = dst.Tables("SOTORDR2").NewRow
                    rowSOTORDR2.Item("ORDR_NO") = ORDR_NO
                    rowSOTORDR2.Item("ORDR_LNO") = Val(rowImportDetails.Item("ORDR_LNO") & String.Empty)
                    rowSOTORDR2.Item("ITEM_CODE") = rowImportDetails.Item("ITEM_CODE") & String.Empty
                    rowSOTORDR2.Item("ITEM_DESC") = rowImportDetails.Item("ITEM_DESC") & String.Empty
                    rowSOTORDR2.Item("ITEM_DESC2") = rowImportDetails.Item("ITEM_DESC2") & String.Empty
                    dst.Tables("SOTORDR2").Rows.Add(rowSOTORDR2)

                    errorCodes = String.Empty
                    If rowSOTORDRX.Item("ITEM_UOM") & String.Empty = "1" Then
                        AddCharNoDups(InvalidUOM, SOTORDR2ErrorCodes)
                    End If
                    SetItemInfo(rowSOTORDR2, errorCodes)
                    AddCharNoDups(errorCodes, SOTORDR2ErrorCodes)

                    If (rowSOTORDR2.Item("ITEM_CODE") & String.Empty).ToString.Trim.Length = 0 Then
                        rowSOTORDR2.Item("ITEM_CODE") = rowImportDetails.Item("ITEM_CODE") & String.Empty
                        rowSOTORDR2.Item("ITEM_DESC") = rowImportDetails.Item("ITEM_DESC") & String.Empty
                        rowSOTORDR2.Item("ITEM_DESC2") = rowImportDetails.Item("ITEM_DESC2") & String.Empty
                    End If

                    rowSOTORDR2.Item("ORDR_UNIT_PRICE") = 0
                    ORDR_QTY = Val(rowImportDetails.Item("ORDR_QTY") & String.Empty)

                    rowSOTORDR2.Item("ORDR_QTY") = ORDR_QTY
                    rowSOTORDR2.Item("ORDR_QTY_OPEN") = ORDR_QTY
                    rowSOTORDR2.Item("ORDR_QTY_ORIG") = ORDR_QTY
                    rowSOTORDR2.Item("ORDR_LINE_STATUS") = "O"
                    rowSOTORDR2.Item("CUST_LINE_REF") = rowImportDetails.Item("CUST_LINE_REF") & String.Empty
                    rowSOTORDR2.Item("ORDR_LINE_SOURCE") = rowImportDetails.Item("ORDR_LINE_SOURCE") & String.Empty
                    rowSOTORDR2.Item("ORDR_LR") = rowImportDetails.Item("ORDR_LR") & String.Empty

                    If ORDR_QTY <= 0 Then
                        Me.AddCharNoDups(QtyOrdered, SOTORDR1ErrorCodes)
                    End If

                    rowSOTORDR2.Item("ORDR_QTY_PICK") = 0
                    rowSOTORDR2.Item("ORDR_QTY_SHIP") = 0
                    rowSOTORDR2.Item("ORDR_QTY_CANC") = 0
                    rowSOTORDR2.Item("ORDR_QTY_BACK") = 0
                    rowSOTORDR2.Item("ORDR_QTY_ONPO") = 0

                    rowSOTORDR2.Item("PATIENT_NAME") = rowImportDetails.Item("PATIENT_NAME") & String.Empty
                    rowSOTORDR2.Item("ORDR_UNIT_PRICE_OVERRIDDEN") = String.Empty
                    rowSOTORDR2.Item("ORDR_QTY_DFCT") = 0
                    rowSOTORDR2.Item("ORDR_UNIT_PRICE_PATIENT") = Val(rowImportDetails.Item("ORDR_UNIT_PRICE_PATIENT") & String.Empty)
                    rowSOTORDR2.Item("HANDLING_CODE") = String.Empty
                    rowSOTORDR2.Item("SAMPLE_IND") = "0"
                    rowSOTORDR2.Item("ORDR_REL_HOLD_CODES") = SOTORDR2ErrorCodes.Trim

                    Me.AddCharNoDups(SOTORDR2ErrorCodes, SOTORDR1ErrorCodes)

                Next

                If Not CreateOrderBillTo(ORDR_NO) Then
                    Me.AddCharNoDups(InvalidSoldTo, SOTORDR1ErrorCodes)
                End If

                If Not CreateOrderShipTo(ORDR_NO, rowSOTORDRX, AlwaysUseImportedShipToAddress OrElse shipToPatient) Then
                    Me.AddCharNoDups(InvalidShipTo, SOTORDR1ErrorCodes)
                End If

                ' Validate DPD Address
                If rowSOTORDR1.Item("ORDR_DPD") = "1" Then
                    If dst.Tables("SOTORDR5").Select("CUST_ADDR_TYPE = 'ST'").Length > 0 Then
                        Dim rowSOTORDR5_ST As DataRow = dst.Tables("SOTORDR5").Select("CUST_ADDR_TYPE = 'ST'")(0)
                        If Not ValidateDPDAddress(rowSOTORDR1, rowSOTORDR5_ST) Then
                            Me.AddCharNoDups(InvalidDPDAddress, SOTORDR1ErrorCodes)
                        End If
                    End If
                End If

                If Not CreateSalesOrderTax(ORDR_NO) Then
                    Me.AddCharNoDups(InvalidSalesTax, SOTORDR1ErrorCodes)
                End If

                SOTORDR1ErrorCodes &= String.Empty
                rowSOTORDR1.Item("ORDR_REL_HOLD_CODES") = SOTORDR1ErrorCodes.Trim

                If (rowSOTORDR1.Item("ORDR_REL_HOLD_CODES") & String.Empty).ToString.Trim.Length = 0 _
                    OrElse (rowSOTORDR1.Item("ORDR_REL_HOLD_CODES") & String.Empty).ToString.Trim = InvalidDPDAddress Then
                    If Not Me.GetSalesOrderUnitPrices(ORDR_NO) Then
                        rowSOTORDR1.Item("ORDR_REL_HOLD_CODES") = rowSOTORDR1.Item("ORDR_REL_HOLD_CODES") & InvalidPricing
                    Else
                        ' Make sure there are no Revenue items where the Order Unit Price is 0
                        Dim revenueAtZero As Boolean = False
                        For Each rowSOTORDR2_P As DataRow In dst.Tables("SOTORDR2").Select("ORDR_UNIT_PRICE = 0")
                            rowICTPCAT1 = baseClass.LookUp("ICTPCAT1", rowSOTORDR2_P.Item("PRICE_CATGY_CODE") & String.Empty)
                            If rowICTPCAT1.Item("PRICE_CATGY_SAMPLE_IND") & String.Empty <> "1" Then
                                rowSOTORDR2_P.Item("ORDR_REL_HOLD_CODES") = (rowSOTORDR2_P.Item("ORDR_REL_HOLD_CODES") & String.Empty).ToString.Trim & RevenueItemNoPrice
                                revenueAtZero = True
                            End If
                        Next
                        If revenueAtZero Then
                            rowSOTORDR1.Item("ORDR_REL_HOLD_CODES") = rowSOTORDR1.Item("ORDR_REL_HOLD_CODES") & RevenueItemNoPrice
                        End If
                    End If
                End If

                If (rowSOTORDR1.Item("ORDR_REL_HOLD_CODES") & String.Empty).ToString.Trim.Length = 0 _
                     OrElse (rowSOTORDR1.Item("ORDR_REL_HOLD_CODES") & String.Empty).ToString.Trim = InvalidDPDAddress Then
                    If Not Me.UpdateSalesOrderTotal(ORDR_NO) Then
                        rowSOTORDR1.Item("ORDR_REL_HOLD_CODES") = rowSOTORDR1.Item("ORDR_REL_HOLD_CODES") & InvalidSalesOrderTotal
                    End If
                End If

                rowSOTORDR1.Item("ORDR_REL_HOLD_CODES") &= DpdOrderWithoutAnAnnualSupply(rowSOTORDR1, dst.Tables("SOTORDR2")) & String.Empty
                rowSOTORDR1.Item("ORDR_REL_HOLD_CODES") &= DpdOrderCODneedsAuthorization(rowSOTORDR1, dst.Tables("SOTORDR2")) & String.Empty

                If (rowSOTORDR1.Item("ORDR_REL_HOLD_CODES") & String.Empty).ToString.Trim.Length > 0 Then
                    rowSOTORDR1.Item("ORDR_STATUS_WEB") = ORDR_STATUS_WEB
                    If ImportErrorNotification.ContainsKey(rowSOTORDR1.Item("ORDR_SOURCE")) Then
                        ImportErrorNotification.Item(rowSOTORDR1.Item("ORDR_SOURCE")) += 1
                    Else
                        ImportErrorNotification.Add(rowSOTORDR1.Item("ORDR_SOURCE"), 1)
                    End If
                End If

                rowSOTORDR1.Item("INIT_DATE") = DateTime.Now + ABSolution.ASCMAIN1.NowTSD
                rowSOTORDR1.Item("LAST_DATE") = rowSOTORDR1.Item("INIT_DATE")

                rowSOTORDR1.Item("INIT_OPER") = ABSolution.ASCMAIN1.USER_ID
                rowSOTORDR1.Item("LAST_OPER") = ABSolution.ASCMAIN1.USER_ID

                If testMode Then RecordLogEntry("Exit CreateSalesOrder: " & ORDR_NO)

                If dst.Tables("SOTORDR5").Select("ORDR_NO = '" & ORDR_NO & "' AND CUST_ADDR_TYPE = 'ST'").Length > 0 Then
                    If rowSOTORDR1.Item("ORDR_DPD") = "1" _
                        AndAlso Not (rowSOTORDR1.Item("ORDR_REL_HOLD_CODES") & String.Empty).ToString.Contains(InvalidShipTo) _
                        AndAlso Not (rowSOTORDR1.Item("ORDR_REL_HOLD_CODES") & String.Empty).ToString.Contains(InvalidSoldTo) Then
                        rowSOTORDR1.Item("PATIENT_NO") = CreateDPDPatientRecord(CUST_CODE, CUST_SHIP_TO_NO, dst.Tables("SOTORDR5").Select("ORDR_NO = '" & ORDR_NO & "' AND CUST_ADDR_TYPE = 'ST'")(0))
                    End If
                End If
                CreateSalesOrder = True

            Catch ex As Exception
                CreateSalesOrder = False
                RecordLogEntry("CreateSalesOrder: " & ex.Message)
                emailErrors(ORDR_SOURCE, 1, ex.Message)
            End Try

        End Function

        Private Function CreateSalesOrderTax(ByVal ORDR_NO As String) As Boolean

            Dim rowSOTORDR1 As DataRow = Nothing

            Try
                If testMode Then RecordLogEntry("Enter CreateSalesOrderTax.")

                Dim CUST_SHIP_TO_ZIP_TAX As String = String.Empty
                Dim CUST_SHIP_TO_STATE As String = String.Empty
                Dim STAX_EXEMPT As String = String.Empty
                Dim STAX_CODE As String = String.Empty
                Dim STAX_RATE As Double = 0

                rowSOTORDR1 = dst.Tables("SOTORDR1").Select("ORDR_NO = '" & ORDR_NO & "'")(0)

                If rowARTCUST2 IsNot Nothing Then
                    CUST_SHIP_TO_ZIP_TAX = (rowARTCUST2.Item("CUST_SHIP_TO_ZIP_CODE") & String.Empty).ToString.Trim
                    CUST_SHIP_TO_STATE = rowARTCUST2.Item("CUST_SHIP_TO_STATE") & String.Empty
                    STAX_EXEMPT = Val(rowARTCUST2.Item("STAX_EXEMPT") & String.Empty).ToString.Trim
                ElseIf rowARTCUST1 IsNot Nothing Then
                    CUST_SHIP_TO_ZIP_TAX = (rowARTCUST1.Item("CUST_ZIP_CODE") & String.Empty).ToString.Trim
                    CUST_SHIP_TO_STATE = rowARTCUST1.Item("CUST_STATE") & String.Empty
                    STAX_EXEMPT = Val(rowARTCUST1.Item("STAX_EXEMPT") & String.Empty).ToString.Trim
                End If

                If CUST_SHIP_TO_ZIP_TAX.Length > 5 Then CUST_SHIP_TO_ZIP_TAX = CUST_SHIP_TO_ZIP_TAX.Substring(0, 5)
                STAX_CODE = CUST_SHIP_TO_STATE

                rowSOTORDR1.Item("CUST_SHIP_TO_STATE") = CUST_SHIP_TO_STATE
                rowSOTORDR1.Item("CUST_SHIP_TO_ZIP_TAX") = CUST_SHIP_TO_ZIP_TAX
                rowSOTORDR1.Item("STAX_EXEMPT") = STAX_EXEMPT

                rowSOTORDR1.Item("CUST_SHIP_TO_STATE") = CUST_SHIP_TO_STATE

                STAX_RATE = TAC.SOCMAIN1.Get_STAX_RATE(baseClass, STAX_CODE, STAX_EXEMPT, CUST_SHIP_TO_ZIP_TAX, STAX_CODE_states)

                rowSOTORDR1.Item("STAX_CODE") = IIf(STAX_CODE_states.Contains(STAX_CODE), STAX_CODE, String.Empty)
                rowSOTORDR1.Item("STAX_RATE") = STAX_RATE

                If testMode Then RecordLogEntry("Exit CreateSalesOrderTax.")
                Return True

            Catch ex As Exception
                RecordLogEntry("CreateSalesOrderTax: (" & ORDR_NO & ") " & ex.Message)
                Return False
            End Try
        End Function

        Private Function GetOrderSalesTaxByState(ByVal rowSOTORDR1 As DataRow, ByVal tblSOTORDR2 As DataTable) As Double

            Dim ORDR_NO As String = String.Empty

            Try
                If testMode Then RecordLogEntry("Enter GetOrderSalesTaxByState.")

                ORDR_NO = rowSOTORDR1.Item("ORDR_NO") & String.Empty
                Dim CUST_CODE As String = rowSOTORDR1.Item("CUST_CODE") & String.Empty
                Dim CUST_SHIP_TO_NO As String = rowSOTORDR1.Item("CUST_SHIP_TO_NO") & String.Empty

                Dim ORDER_TOTAL As Double = Val(tblSOTORDR2.Compute("SUM(ORDR_LNO_EXT)", "ORDR_NO = '" & ORDR_NO & "'") & String.Empty)
                Dim taxableAmount As Double = ORDER_TOTAL
                Dim rowARTSTAX1 As DataRow = baseClass.LookUp("ARTSTAX1", rowSOTORDR1.Item("STAX_CODE") & String.Empty)

                If rowARTSTAX1 IsNot Nothing Then
                    If rowARTSTAX1.Item("STAX_ON_FRT") & String.Empty = "1" Then
                        taxableAmount += Val(rowSOTORDR1.Item("ORDR_FREIGHT") & String.Empty)
                    End If

                    If rowARTSTAX1.Item("STAX_ON_MISC") & String.Empty = "1" Then
                        taxableAmount += (Val(rowSOTORDR1.Item("ORDR_SAMPLE_SURCHARGE") & String.Empty) + Val(rowSOTORDR1.Item("ORDR_MISC_CHG_AMT") & String.Empty))
                    End If
                Else
                    rowSOTORDR1.Item("STAX_CODE") = String.Empty
                End If

                If testMode Then RecordLogEntry("Exit GetOrderSalesTaxByState.")
                Return Math.Round((taxableAmount * rowSOTORDR1.Item("STAX_RATE")) / 100, 2, MidpointRounding.AwayFromZero)

            Catch ex As Exception
                RecordLogEntry("GetOrderSalesTaxByState: (" & ORDR_NO & ")" & ex.Message)
                Return -1
            End Try


        End Function

        Private Function GetSalesOrderUnitPrices(ByVal ORDR_NO As String) As Boolean

            Dim rowSOTORDR1 As DataRow = Nothing

            Try
                If testMode Then RecordLogEntry("Enter GetSalesOrderUnitPrices.")

                rowSOTORDR1 = dst.Tables("SOTORDR1").Select("ORDR_NO = '" & ORDR_NO & "'")(0)
                If Not Me.TestAuthorizationsAndBlocks(rowSOTORDR1) Then
                    Return False
                End If

                clsSOCORDR1.AffiliateFreeShipping()
                clsSOCORDR1.Price_and_Qty(False, True)

                'Dim tblSubs As DataTable = clsSOCORDR1.ItemSubsitiutions
                'If tblSubs.Rows.Count > 0 Then
                '    For Each rowSubs As DataRow In tblSubs.Select("", "ORDR_LNO")

                '        Dim rowSOTORDR2_ORIG As DataRow = dst.Tables("SOTORDR2").Select("ORDR_LNO = " & rowSubs.Item("ORDR_LNO_ORIG"))(0)
                '        Dim rowSOTORDR2 As DataRow = dst.Tables("SOTORDR2").NewRow
                '        rowSOTORDR2.Item("ORDR_NO") = ORDR_NO
                '        rowSOTORDR2.Item("ORDR_LNO") = rowSubs.Item("ORDR_LNO")
                '        rowSOTORDR2.Item("ITEM_CODE") = rowSubs.Item("ITEM_CODE")
                '        rowSOTORDR2.Item("ORDR_QTY") = Val(rowSubs.Item("ORDR_QTY") & String.Empty)
                '        rowSOTORDR2.Item("ORDR_UNIT_PRICE") = Val(rowSubs.Item("ORDR_UNIT_PRICE") & String.Empty)
                '        rowSOTORDR2.Item("ORDR_UNIT_PRICE_PATIENT") = Val(rowSubs.Item("ORDR_UNIT_PRICE_PATIENT") & String.Empty)
                '        rowSOTORDR2.Item("ORDR_UNIT_PRICE_OVERRIDDEN") = rowSubs.Item("ORDR_UNIT_PRICE_OVERRIDDEN") & String.Empty
                '        rowSOTORDR2.Item("ORDR_UNIT_PRICE_MANUAL") = Val(rowSubs.Item("ORDR_UNIT_PRICE") & String.Empty)
                '        rowSOTORDR2.Item("ORDR_UNIT_PRICE_SOURCE") = rowSubs.Item("ORDR_UNIT_PRICE_SOURCE") & String.Empty
                '        rowSOTORDR2.Item("ITEM_UOM") = rowSubs.Item("ITEM_UOM") & String.Empty
                '        rowSOTORDR2.Item("PATIENT_NAME") = rowSOTORDR2_ORIG.Item("PATIENT_NAME") & String.Empty
                '        rowSOTORDR2.Item("ORDR_LR") = rowSOTORDR2_ORIG.Item("ORDR_LR") & String.Empty
                '        rowSOTORDR2.Item("PATIENT_GROUP") = rowSOTORDR2_ORIG.Item("PATIENT_GROUP") & String.Empty
                '        dst.Tables("SOTORDR2").Rows.Add(rowSOTORDR2)

                '        rowSOTORDR2_ORIG.Item("ORDR_QTY") = 0
                '    Next

                '    ' Reprice the Sales Order after Substitutions
                '    clsSOCORDR1.Price_and_Qty(False, True)
                'End If

                ' Added on 1/22/2009 as per walter
                If clsSOCORDR1.SHIP_VIA_CODE_switch_to.Trim.Length > 0 Then
                    If (rowSOTORDR1.Item("ORDR_LOCK_SHIP_VIA") & String.Empty) <> "1" Then
                        rowSOTORDR1.Item("SHIP_VIA_CODE") = clsSOCORDR1.SHIP_VIA_CODE_switch_to.Trim
                    End If
                End If

                If clsSOCORDR1.ORDR_NO_FREIGHT & String.Empty = "1" _
                    AndAlso rowSOTORDR1.Item("ORDR_NO_FREIGHT") & String.Empty <> "1" Then
                    rowSOTORDR1.Item("ORDR_NO_FREIGHT") = "1"
                    rowSOTORDR1.Item("REASON_CODE_NO_FRT") = clsSOCORDR1.REASON_CODE_NO_FRT
                End If

                If testMode Then RecordLogEntry("Exit GetSalesOrderUnitPrices.")
                Return True

            Catch ex As Exception
                RecordLogEntry("GetSalesOrderUnitPrices: (" & ORDR_NO & ") - " & ex.Message)
                Return False
            End Try

        End Function

        ''' <summary>
        ''' Returns the sample for an item, if one exists
        ''' </summary>
        ''' <param name="ITEM_CODE"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Private Function GetSampleItem(ByVal ITEM_CODE As String) As DataRow

            Dim sql As String = String.Empty
            Dim sample As DataRow = Nothing

            Try
                sql = "  SELECT * FROM ICTITEM1"
                sql &= " WHERE PRICE_CATGY_CODE = (SELECT PRICE_CATGY_CODE_SAMPLE FROM ICTPCAT1 WHERE PRICE_CATGY_CODE = (SELECT PRICE_CATGY_CODE FROM ICTITEM1 WHERE ITEM_CODE = :PARM1 )) "
                sql &= " AND (NVL(ITEM_BASE_CURVE, 0), NVL(ITEM_SPHERE_POWER, 0), NVL(ITEM_DIAMETER, 0), NVL(ITEM_AXIS, 0), NVL(ITEM_CYLINDER, 0), NVL(ITEM_ADD_POWER, 0)) IN"
                sql &= " (SELECT NVL(ITEM_BASE_CURVE, 0), NVL(ITEM_SPHERE_POWER, 0), NVL(ITEM_DIAMETER, 0), NVL(ITEM_AXIS, 0), NVL(ITEM_CYLINDER, 0), NVL(ITEM_ADD_POWER, 0)"
                sql &= " FROM ICTITEM1 WHERE ITEM_CODE = :PARM2)"
                sql &= " AND ITEM_STATUS = 'A'"

                sample = ABSolution.ASCDATA1.GetDataRow(sql, "VV", New Object() {ITEM_CODE, ITEM_CODE})
            Catch ex As Exception
            End Try

            Return sample

        End Function

        ''' <summary>
        ''' Record sales order event
        ''' </summary>
        ''' <param name="ORDR_NO"></param>
        ''' <param name="EVENT_DESC"></param>
        ''' <remarks></remarks>
        Private Sub Record_Event(ByVal ORDR_NO As String, ByVal EVENT_DESC As String)

            Try

                If testMode Then RecordLogEntry("Enter Record_Event.")

                Dim row As DataRow = dst.Tables("SOTORDRE").NewRow
                row.Item("ORDR_NO") = ORDR_NO
                row.Item("INIT_DATE") = DateTime.Now
                row.Item("INIT_OPER") = ABSolution.ASCMAIN1.USER_ID
                row.Item("EVENT_DESC") = EVENT_DESC
                dst.Tables("SOTORDRE").Rows.Add(row)

                If testMode Then RecordLogEntry("Exit Record_Event.")

            Catch ex As Exception
                RecordLogEntry("Record_Event: (" & ORDR_NO & ") " & ex.Message)
            End Try
        End Sub

        ''' <summary>
        ''' Sets Order Header Information based on the Bill To Customer Attributes
        ''' </summary>
        ''' <param name="CUST_CODE"></param>
        ''' <param name="CUST_SHIP_TO_NO"></param>
        ''' <param name="rowSOTORDR1"></param>
        ''' <param name="errorCodes"></param>
        ''' <remarks></remarks>
        Private Function SetBillToAttributes(ByVal CUST_CODE As String, ByVal CUST_SHIP_TO_NO As String, ByRef rowSOTORDR1 As DataRow, ByRef errorCodes As String) As Boolean

            Try

                If testMode Then RecordLogEntry("Enter SetBillToAttributes.")

                Dim STAX_CODE As String = String.Empty
                Dim TERM_CODE As String = String.Empty
                Dim SHIP_VIA_CODE_FIELD As String = String.Empty
                Dim STAX_EXEMPT As String = String.Empty
                Dim STAX_RATE As String = String.Empty
                Dim zipCode As String = String.Empty

                Dim ORDR_NO As String = rowSOTORDR1.Item("ORDR_NO") & String.Empty
                If errorCodes Is Nothing Then errorCodes = String.Empty

                ' Keep data from the imported data to help assist with any invalid codes
                'rowSOTORDR1.Item("CUST_NAME") = String.Empty
                'rowSOTORDR1.Item("CUST_BILL_TO_CUST") = String.Empty

                If (rowSOTORDR1.Item("ORDR_LOCK_SHIP_VIA") & String.Empty) <> "1" Then
                    'rowSOTORDR1.Item("SHIP_VIA_CODE") = String.Empty
                End If

                rowSOTORDR1.Item("POST_CODE") = String.Empty
                rowSOTORDR1.Item("TERM_CODE") = String.Empty
                rowSOTORDR1.Item("SREP_CODE") = String.Empty
                rowSOTORDR1.Item("ORDR_SHIP_COMPLETE") = String.Empty
                rowSOTORDR1.Item("STAX_CODE") = String.Empty

                If rowARTCUST1 IsNot Nothing Then
                    rowSOTORDR1.Item("CUST_CODE") = rowARTCUST1.Item("CUST_CODE") & String.Empty
                    rowSOTORDR1.Item("CUST_NAME") = rowARTCUST1.Item("CUST_NAME") & String.Empty

                    rowSOTORDR1.Item("CUST_BILL_TO_CUST") = rowARTCUST1.Item("CUST_BILL_TO_CUST") & String.Empty
                    If (rowARTCUST1.Item("CUST_BILL_TO_CUST") & String.Empty).ToString.Trim.Length = 0 Then
                        rowSOTORDR1.Item("CUST_BILL_TO_CUST") = rowARTCUST1.Item("CUST_CODE") & String.Empty
                    End If

                    TERM_CODE = rowARTCUST1.Item("TERM_CODE") & String.Empty
                    rowTATTERM1 = baseClass.LookUp("TATTERM1", TERM_CODE)
                    If rowTATTERM1 IsNot Nothing Then
                        rowSOTORDR1.Item("TERM_CODE") = TERM_CODE
                    Else
                        AddCharNoDups(InvalidTermsCode, errorCodes)
                    End If

                    rowSOTORDR1.Item("POST_CODE") = rowARTCUST1.Item("POST_CODE") & String.Empty
                    rowSOTORDR1.Item("SREP_CODE") = rowARTCUST1.Item("SREP_CODE") & String.Empty
                    rowSOTORDR1.Item("ORDR_SHIP_COMPLETE") = rowARTCUST1.Item("CUST_SHIP_COMPLETE") & String.Empty
                    rowSOTORDR1.Item("ORDR_NO_SAMPLE_SURCHARGE") = rowARTCUST1.Item("NO_SAMPLE_SURCHARGE") & String.Empty
                    rowSOTORDR1.Item("ORDR_NO_SAMPLE_HANDLING_FEE") = rowARTCUST1.Item("NO_SAMPLE_HANDLING_FEE") & String.Empty
                Else
                    If (rowSOTORDR1.Item("CUST_NAME") & String.Empty).ToString.Trim.Length = 0 Then
                        rowSOTORDR1.Item("CUST_NAME") = "Unknown Customer"
                    End If
                    rowSOTORDR1.Item("CUST_BILL_TO_CUST") = String.Empty
                    Me.AddCharNoDups(InvalidSoldTo, errorCodes)
                End If

                ' Always use the ship via code sent by the customer
                If rowSOTORDR1.Item("SHIP_VIA_CODE") & String.Empty = String.Empty Then
                    If rowARTCUST2 IsNot Nothing Then
                        If rowSOTORDR1.Item("ORDR_DPD") & String.Empty <> "1" Then
                            If (rowARTCUST2.Item("CUST_SHIP_TO_SHIP_VIA_CODE") & String.Empty).ToString.Length > 0 Then
                                rowSOTORDR1.Item("SHIP_VIA_CODE") = rowARTCUST2.Item("CUST_SHIP_TO_SHIP_VIA_CODE") & String.Empty
                            End If
                        Else
                            If (rowARTCUST2.Item("CUST_SHIP_TO_SHIP_VIA_CODE_DPD") & String.Empty).ToString.Length > 0 Then
                                rowSOTORDR1.Item("SHIP_VIA_CODE") = rowARTCUST2.Item("CUST_SHIP_TO_SHIP_VIA_CODE_DPD") & String.Empty
                            End If
                        End If
                    End If

                    ' If the ship via is not set then look at freight contract
                    If rowARTCUST3 IsNot Nothing AndAlso rowSOTORDR1.Item("SHIP_VIA_CODE") & String.Empty = String.Empty Then
                        If rowSOTORDR1.Item("ORDR_DPD") & String.Empty = "1" Then
                            rowSOTORDR1.Item("SHIP_VIA_CODE") = rowARTCUST3.Item("SHIP_VIA_CODE_DPD") & String.Empty
                            If rowSOTORDR1.Item("SHIP_VIA_CODE") & String.Empty = String.Empty Then
                                rowSOTORDR1.Item("SHIP_VIA_CODE") = DpdDefaultShipViaCode
                            End If
                        Else
                            rowSOTORDR1.Item("SHIP_VIA_CODE") = rowARTCUST3.Item("SHIP_VIA_CODE") & String.Empty
                        End If
                    End If
                End If

                rowSOTSVIA1 = baseClass.LookUp("SOTSVIA1", rowSOTORDR1.Item("SHIP_VIA_CODE") & String.Empty)

                If rowARTCUST3 IsNot Nothing Then
                    rowSOTORDR1.Item("FRT_CONT_NO") = rowARTCUST3.Item("FRT_CONT_NO") & String.Empty
                End If

                If testMode Then RecordLogEntry("Exit SetBillToAttributes.")
                Return True

            Catch ex As Exception
                RecordLogEntry("SetBillToAttributes: (" & rowSOTORDR1.Item("ORDR_NO") & ") " & ex.Message)
                Return False
            End Try

        End Function

        ''' <summary>
        ''' Overrides Order Settings for DPD orders
        ''' </summary>
        ''' <param name="rowSOTORDR1"></param>
        ''' <param name="errorCodes"></param>
        ''' <remarks></remarks>
        Private Function SetDPDShipViaSettings(ByRef rowSOTORDR1 As DataRow, ByRef errorCodes As String, ByVal ORDR_LINE_SOURCE As String) As Boolean

            Try

                If testMode Then RecordLogEntry("Enter SetDPDShipViaSettings.")

                If (rowSOTORDR1.Item("SHIP_VIA_CODE") & String.Empty) <> String.Empty Then Return True
                If (rowSOTORDR1.Item("ORDR_DPD") & String.Empty) <> "1" Then Return True
                'If (rowSOTORDR1.Item("ORDR_LOCK_SHIP_VIA") & String.Empty) = "1" Then Return True

                Dim SHIP_VIA_CODE_DPD As String = String.Empty
                Dim originalSHIP_VIA_CODE_DPD As String = rowSOTORDR1.Item("SHIP_VIA_CODE") & String.Empty

                ' Look at Ship-To First
                If rowARTCUST2 IsNot Nothing Then
                    SHIP_VIA_CODE_DPD = (rowARTCUST2.Item("CUST_SHIP_TO_SHIP_VIA_CODE_DPD") & String.Empty).ToString.Trim
                End If

                ' If not Ship-To Ship Via then look at Customer Contract
                If SHIP_VIA_CODE_DPD.Length = 0 Then
                    If rowARTCUST3 IsNot Nothing Then
                        SHIP_VIA_CODE_DPD = (rowARTCUST3.Item("SHIP_VIA_CODE_DPD") & String.Empty).ToString.Trim
                    End If
                End If

                ' If not Customer DPD contract then use system default
                If SHIP_VIA_CODE_DPD.Length = 0 Then
                    Select Case ORDR_LINE_SOURCE
                        Case "Y" ' eyeconic
                            SHIP_VIA_CODE_DPD = originalSHIP_VIA_CODE_DPD

                        Case Else
                            SHIP_VIA_CODE_DPD = DpdDefaultShipViaCode
                    End Select

                    If SHIP_VIA_CODE_DPD.Length = 0 Then
                        SHIP_VIA_CODE_DPD = DpdDefaultShipViaCode
                    End If
                End If

                rowSOTSVIA1 = baseClass.LookUp("SOTSVIA1", SHIP_VIA_CODE_DPD)

                If rowSOTSVIA1 IsNot Nothing Then
                    rowSOTORDR1.Item("SHIP_VIA_CODE") = rowSOTSVIA1.Item("SHIP_VIA_CODE") & String.Empty
                    errorCodes = Replace(errorCodes, InvalidShipVia, String.Empty)
                End If

                If testMode Then RecordLogEntry("Exit SetDPDShipViaSettings.")
                Return True
            Catch ex As Exception
                RecordLogEntry("SetDPDShipViaSettings: (" & rowSOTORDR1.Item("ORDR_NO") & ") " & ex.Message)
                Return False
            End Try

        End Function

        ''' <summary>
        ''' Sets Order Detail Information based on the Item's Attributes
        ''' </summary>
        ''' <param name="rowSOTORDR2"></param>
        ''' <param name="errorCodes"></param>
        ''' <remarks></remarks>
        Private Sub SetItemInfo(ByRef rowSOTORDR2 As DataRow, ByRef errorCodes As String)
            Try

                If testMode Then RecordLogEntry("Enter SetItemInfo.")

                Dim ITEM_CODE As String = (rowSOTORDR2.Item("ITEM_CODE") & String.Empty).ToString.Trim
                Dim ITEM_DESC2 As String = rowSOTORDR2.Item("ITEM_DESC2") & String.Empty
                Dim ITEM_DESC2_X() As String = ITEM_DESC2.Split("/")

                rowICTITEM1 = Nothing

                If ITEM_CODE.Length > 0 Then

                    ' See if the item exists in the item master
                    If ITEM_CODE.Length > 0 Then
                        rowICTITEM1 = baseClass.LookUp("ICTITEM1", ITEM_CODE)
                    End If

                    ' If the item is not in the Item master then look at the catalogue
                    If rowICTITEM1 Is Nothing Then
                        rowICTITEM1 = baseClass.LookUp("ICTCATL1", ITEM_CODE)
                    End If

                    ' this may be a UPC Code
                    If rowICTITEM1 Is Nothing Then
                        rowICTITEM1 = ABSolution.ASCDATA1.GetDataRow("SELECT * FROM ICTITEM1 WHERE ITEM_UPC_CODE = :PARM1", "V", New Object() {ITEM_CODE})
                    End If

                End If

                ' If the item is not found but the desc2 has it attributes then try to get it by attributes
                If rowICTITEM1 Is Nothing AndAlso ITEM_DESC2_X.Length = 8 Then

                    Dim sql As String = "SELECT * FROM ICTCATL1  "
                    sql &= " WHERE PRICE_CATGY_CODE = :PARM1"
                    sql &= " AND NVL(ITEM_BASE_CURVE, 0) = :PARM2"
                    sql &= " AND NVL(ITEM_SPHERE_POWER, 0) = :PARM3"
                    sql &= " AND NVL(ITEM_CYLINDER, 0) = :PARM4"
                    sql &= " AND NVL(ITEM_AXIS, 0) = :PARM5"
                    sql &= " AND NVL(ITEM_DIAMETER, 0) = :PARM6"
                    sql &= " AND ITEM_ADD_POWER = :PARM7"
                    sql &= " AND NVL(ITEM_COLOR, '@@') = :PARM8"

                    If ITEM_DESC2_X(7).Length = 0 Then
                        ITEM_DESC2_X(7) = "@@"
                    End If

                    If ITEM_DESC2_X(6).Length = 0 Then
                        ITEM_DESC2_X(6) = "0.00"
                    End If

                    Try
                        rowICTITEM1 = ABSolution.ASCDATA1.GetDataRow(sql, "VNNNNNVV", New Object() {ITEM_DESC2_X(0), _
                            Val(ITEM_DESC2_X(1)), Val(ITEM_DESC2_X(2)), _
                            Val(ITEM_DESC2_X(3)), Val(ITEM_DESC2_X(4)), _
                            Val(ITEM_DESC2_X(5)), ITEM_DESC2_X(6), ITEM_DESC2_X(7)})
                    Catch ex As Exception
                        rowICTITEM1 = Nothing
                    End Try

                End If

                If rowICTITEM1 IsNot Nothing Then
                    Dim itemToCreate As String = rowICTITEM1.Item("ITEM_CODE")
                    rowICTITEM1 = baseClass.LookUp("ICTITEM1", itemToCreate)
                    If rowICTITEM1 Is Nothing Then
                        Try
                            TAC.ICCMAIN1.Create_Item_from_Catalog(itemToCreate)
                        Catch ex As Exception
                            ' Nothing
                        End Try
                        rowICTITEM1 = baseClass.LookUp("ICTITEM1", itemToCreate)
                    End If
                    ITEM_CODE = itemToCreate
                End If

                If rowICTITEM1 Is Nothing Then
                    Me.AddCharNoDups(InvalidItem, errorCodes)
                Else
                    rowSOTORDR2.Item("ITEM_CODE") = rowICTITEM1.Item("ITEM_CODE") & String.Empty
                    rowSOTORDR2.Item("ITEM_DESC") = rowICTITEM1.Item("ITEM_DESC") & String.Empty
                    rowSOTORDR2.Item("ITEM_DESC2") = rowICTITEM1.Item("ITEM_DESC2") & String.Empty
                    rowSOTORDR2.Item("ITEM_UOM") = "EA"
                    rowSOTORDR2.Item("PRICE_CATGY_CODE") = rowICTITEM1.Item("PRICE_CATGY_CODE") & String.Empty

                    If rowICTITEM1.Item("ITEM_ORDER_CODE") & String.Empty = "X" OrElse rowICTITEM1.Item("ITEM_STATUS") & String.Empty = "I" Then
                        Me.AddCharNoDups(FrozenInactiveItem, errorCodes)
                    End If
                End If

                If testMode Then RecordLogEntry("Exit SetItemInfo.")

            Catch ex As Exception
                RecordLogEntry("SetItemInfo: (" & rowSOTORDR2.Item("ORDR_NO") & ") " & ex.Message)
                Me.AddCharNoDups(InvalidItem, errorCodes)
            End Try

        End Sub

        ''' <summary>
        ''' Ship Tos have Overrides to the Customer Master data.
        ''' </summary>
        ''' <param name="rowSOTORDR1"></param>
        ''' <param name="errorCodes"></param>
        ''' <remarks></remarks>
        Private Function SetShipToAttributes(ByRef rowSOTORDR1 As DataRow, ByRef errorCodes As String) As Boolean

            Try

                If testMode Then RecordLogEntry("Enter SetShipToAttributes.")

                Dim CUST_CODE As String = rowSOTORDR1.Item("CUST_CODE")
                Dim CUST_SHIP_TO_NO As String = rowSOTORDR1.Item("CUST_SHIP_TO_NO")

                rowSOTORDR1.Item("CUST_SHIP_TO_NO") = String.Empty
                rowSOTORDR1.Item("CUST_SHIP_TO_NAME") = String.Empty

                If rowARTCUST2 IsNot Nothing Then

                    rowSOTORDR1.Item("CUST_SHIP_TO_NO") = CUST_SHIP_TO_NO
                    rowSOTORDR1.Item("CUST_SHIP_TO_NAME") = rowARTCUST2.Item("CUST_SHIP_TO_NAME") & String.Empty

                    ' Ship to's have overrides for Tax code, Ship Via Code and Term Code
                    If (rowARTCUST2.Item("CUST_SHIP_TO_STAX_CODE") & String.Empty).ToString.Length > 0 Then
                        rowSOTORDR1.Item("STAX_CODE") = rowARTCUST2.Item("CUST_SHIP_TO_STAX_CODE") & String.Empty
                    End If

                    If rowSOTORDR1.Item("SHIP_VIA_CODE") & String.Empty = String.Empty Then
                        If rowSOTORDR1.Item("ORDR_DPD") & String.Empty <> "1" Then
                            If (rowARTCUST2.Item("CUST_SHIP_TO_SHIP_VIA_CODE") & String.Empty).ToString.Length > 0 Then
                                rowSOTORDR1.Item("SHIP_VIA_CODE") = rowARTCUST2.Item("CUST_SHIP_TO_SHIP_VIA_CODE") & String.Empty
                            End If
                        Else
                            If (rowARTCUST2.Item("CUST_SHIP_TO_SHIP_VIA_CODE_DPD") & String.Empty).ToString.Length > 0 Then
                                rowSOTORDR1.Item("SHIP_VIA_CODE") = rowARTCUST2.Item("CUST_SHIP_TO_SHIP_VIA_CODE_DPD") & String.Empty
                            End If
                        End If
                    End If

                    If rowARTCUST2.Item("CUST_SHIP_TO_STATUS") & String.Empty = "C" Then
                        Me.AddCharNoDups(ShipToClosed, errorCodes)
                    ElseIf rowARTCUST2.Item("CUST_SHIP_TO_ORDER_BLOCK") & String.Empty = "1" Then
                        Me.AddCharNoDups(ShipToOrderBlocked, errorCodes)
                    End If

                    rowSOTORDR1.Item("CUST_SHIP_TO_STATE") = rowARTCUST2.Item("CUST_SHIP_TO_STATE") & String.Empty
                    rowSOTORDR1.Item("CUST_SHIP_TO_ZIP_TAX") = If((rowARTCUST2.Item("CUST_SHIP_TO_ZIP_CODE") & String.Empty).ToString.Length > 5, (rowARTCUST2.Item("CUST_SHIP_TO_ZIP_CODE") & String.Empty).ToString().Substring(0, 5), rowARTCUST2.Item("CUST_SHIP_TO_ZIP_CODE") & String.Empty)

                ElseIf CUST_SHIP_TO_NO.Trim.Length > 0 Then
                    rowSOTORDR1.Item("CUST_SHIP_TO_NO") = CUST_SHIP_TO_NO
                    rowSOTORDR1.Item("CUST_SHIP_TO_NAME") = "Unknown"
                    Me.AddCharNoDups(InvalidShipTo, errorCodes)
                    Return False
                End If

                If testMode Then RecordLogEntry("Exit SetShipToAttributes.")
                Return True
            Catch ex As Exception
                RecordLogEntry("SetShipToAttributes: (" & rowSOTORDR1.Item("ORDR_NO") & ") " & ex.Message)
                Return False
            End Try

        End Function

        Private Function TestAuthorizationsAndBlocks(ByRef rowSOTORDR1 As DataRow) As Boolean

            If clsSOCORDR1 Is Nothing Then Return True

            Try
                If testMode Then RecordLogEntry("Enter TestAuthorizationsAndBlocks.")

                Dim CUST_CODE As String = (rowSOTORDR1.Item("CUST_CODE") & String.Empty).ToString.Trim
                Dim CUST_SHIP_TO_NO As String = (rowSOTORDR1.Item("CUST_SHIP_TO_NO") & String.Empty).ToString.Trim
                Dim ITEM_LIST As String = String.Empty
                Dim ORDR_NO As String = rowSOTORDR1.Item("ORDR_NO")
                Dim ORDR_REL_HOLD_CODES As String = String.Empty

                ' Remove Header Error Code
                ORDR_REL_HOLD_CODES = rowSOTORDR1.Item("ORDR_REL_HOLD_CODES") & String.Empty
                ORDR_REL_HOLD_CODES = ORDR_REL_HOLD_CODES.Replace(ItemAuthorizationError, "") & String.Empty
                rowSOTORDR1.Item("ORDR_REL_HOLD_CODES") = ORDR_REL_HOLD_CODES

                ORDR_REL_HOLD_CODES = String.Empty

                For Each rowSOTORDR2 As DataRow In dst.Tables("SOTORDR2").Select("ISNULL(ORDR_REL_HOLD_CODES,'@@@') <> '@@@'")
                    ITEM_LIST = "'" & rowSOTORDR2.Item("ITEM_CODE") & "'"
                    ORDR_REL_HOLD_CODES = rowSOTORDR2.Item("ORDR_REL_HOLD_CODES") & String.Empty

                    ' Remove Detail Error Code
                    ORDR_REL_HOLD_CODES = ORDR_REL_HOLD_CODES.Replace(ItemAuthorizationError, "") & String.Empty

                    Dim errors As String = clsSOCORDR1.TestAuthorizationsAndBlocks(CUST_CODE, CUST_SHIP_TO_NO, ITEM_LIST, False)

                    errors = errors.Trim
                    If errors.Length = 0 Then Continue For

                    ORDR_REL_HOLD_CODES &= ItemAuthorizationError

                    For Each authError As String In errors.Split(vbCr)
                        authError = authError.Trim
                        If authError.Length > 0 Then
                            CreateOrderErrorRecord(ORDR_NO, 0, ItemAuthorizationError, authError)
                        End If
                    Next

                    ' Place the error on the header
                    If Not (rowSOTORDR1.Item("ORDR_REL_HOLD_CODES") & String.Empty).ToString.Contains(ItemAuthorizationError) Then
                        rowSOTORDR1.Item("ORDR_REL_HOLD_CODES") &= ItemAuthorizationError
                    End If

                    rowSOTORDR2.Item("ORDR_REL_HOLD_CODES") = ORDR_REL_HOLD_CODES
                Next

                If testMode Then RecordLogEntry("Exit TestAuthorizationsAndBlocks.")
                Return True

            Catch ex As Exception
                RecordLogEntry("TestAuthorizationsAndBlocks: (" & rowSOTORDR1.Item("ORDR_NO") & ") " & ex.Message)
                Return False
            End Try

        End Function

        ''' <summary>
        ''' Truncates a fields value if the length is longer the the max length of the field
        ''' </summary>
        ''' <param name="fieldValue"></param>
        ''' <param name="TableName"></param>
        ''' <param name="FieldName"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Private Function TruncateField(ByVal fieldValue As String, ByVal TableName As String, ByVal FieldName As String) As String

            Try
                If testMode Then RecordLogEntry("Enter TruncateField.")

                Dim rValue As String = fieldValue

                If Not dst.Tables.Contains(TableName) Then
                    Return rValue
                End If

                If Not dst.Tables(TableName).Columns.Contains(FieldName) Then
                    Return rValue
                End If

                Dim maxLength As Int16 = 0
                maxLength = dst.Tables(TableName).Columns(FieldName).MaxLength

                If rValue.Length > maxLength Then
                    rValue = rValue.Substring(0, maxLength).Trim
                End If

                If testMode Then RecordLogEntry("Exit TruncateField.")
                Return rValue
            Catch ex As Exception
                RecordLogEntry("TruncateField: " & ex.Message)
                Return fieldValue & String.Empty
            End Try
        End Function

        Private Function UpdateCustomerShipTo(ByVal CUST_CODE As String, ByVal CUST_SHIP_TO_NO As String _
                                           , ByRef rowSOTORDRX As DataRow, ByRef rowARTCUST2 As DataRow) As Boolean

            Try

                If testMode Then RecordLogEntry("Enter UpdateCustomerShipTo: " & CUST_CODE)

                If rowARTCUST2 Is Nothing Then Exit Function
                If rowSOTORDRX Is Nothing Then Exit Function

                Dim sql As String = ""

                Dim CUST_SHIP_TO_NAME As String = StrConv((rowSOTORDRX.Item("CUST_SHIP_TO_NAME") & String.Empty).ToString.Replace("'", "").Trim, VbStrConv.ProperCase)
                Dim CUST_SHIP_TO_ADDR1 As String = StrConv((rowSOTORDRX.Item("CUST_SHIP_TO_ADDR1") & String.Empty).ToString.Replace("'", "").Trim, VbStrConv.ProperCase)
                Dim CUST_SHIP_TO_ADDR2 As String = StrConv((rowSOTORDRX.Item("CUST_SHIP_TO_ADDR2") & String.Empty).ToString.Replace("'", "").Trim, VbStrConv.ProperCase)
                Dim CUST_SHIP_TO_ADDR3 As String = StrConv((rowSOTORDRX.Item("CUST_SHIP_TO_ADDR3") & String.Empty).ToString.Replace("'", "").Trim, VbStrConv.ProperCase)
                Dim CUST_SHIP_TO_CITY As String = StrConv((rowSOTORDRX.Item("CUST_SHIP_TO_CITY") & String.Empty).ToString.Replace("'", "").Trim, VbStrConv.ProperCase)
                Dim CUST_SHIP_TO_STATE As String = (rowSOTORDRX.Item("CUST_SHIP_TO_STATE") & String.Empty).ToString.Replace("'", "").Trim
                Dim CUST_SHIP_TO_ZIP_CODE As String = (rowSOTORDRX.Item("CUST_SHIP_TO_ZIP_CODE") & String.Empty).ToString.Replace("'", "").Trim
                Dim CUST_SHIP_TO_COUNTRY As String = (rowSOTORDRX.Item("CUST_SHIP_TO_COUNTRY") & String.Empty).ToString.Replace("'", "").Trim

                Dim CUST_SHIP_TO_PHONE As String = String.Empty
                For Each chPhone As Char In rowSOTORDRX.Item("CUST_SHIP_TO_PHONE") & String.Empty
                    If Char.IsDigit(chPhone) Then
                        CUST_SHIP_TO_PHONE &= chPhone
                    End If
                Next
                CUST_SHIP_TO_PHONE = TruncateField(CUST_SHIP_TO_PHONE, "ARTCUST2", "CUST_SHIP_TO_PHONE")

                Dim CUST_SHIP_TO_URL As String = TruncateField(rowSOTORDRX.Item("OFFICE_WEBSITE") & String.Empty, "ARTCUST2", "CUST_SHIP_TO_URL").ToLower

                CUST_SHIP_TO_NAME = TruncateField(CUST_SHIP_TO_NAME, "ARTCUST2", "CUST_SHIP_TO_NAME")
                CUST_SHIP_TO_ADDR1 = TruncateField(CUST_SHIP_TO_ADDR1, "ARTCUST2", "CUST_SHIP_TO_ADDR1")
                CUST_SHIP_TO_ADDR2 = TruncateField(CUST_SHIP_TO_ADDR2, "ARTCUST2", "CUST_SHIP_TO_ADDR2")
                CUST_SHIP_TO_ADDR3 = TruncateField(CUST_SHIP_TO_ADDR3, "ARTCUST2", "CUST_SHIP_TO_ADDR3")
                CUST_SHIP_TO_CITY = TruncateField(CUST_SHIP_TO_CITY, "ARTCUST2", "CUST_SHIP_TO_CITY")
                CUST_SHIP_TO_STATE = TruncateField(CUST_SHIP_TO_STATE, "ARTCUST2", "CUST_SHIP_TO_STATE")
                CUST_SHIP_TO_ZIP_CODE = TruncateField(CUST_SHIP_TO_ZIP_CODE, "ARTCUST2", "CUST_SHIP_TO_ZIP_CODE")
                CUST_SHIP_TO_COUNTRY = TruncateField(CUST_SHIP_TO_COUNTRY, "ARTCUST2", "CUST_SHIP_TO_COUNTRY")
                CUST_SHIP_TO_PHONE = TruncateField(CUST_SHIP_TO_PHONE, "ARTCUST2", "CUST_SHIP_TO_PHONE")
                CUST_SHIP_TO_URL = TruncateField(CUST_SHIP_TO_URL, "ARTCUST2", "CUST_SHIP_TO_URL")


                If rowARTCUST2.Item("CUST_SHIP_TO_PHONE") & String.Empty <> CUST_SHIP_TO_PHONE _
                    OrElse rowARTCUST2.Item("CUST_SHIP_TO_URL") & String.Empty <> CUST_SHIP_TO_URL _
                    OrElse rowARTCUST2.Item("CUST_SHIP_TO_NAME") & String.Empty <> CUST_SHIP_TO_NAME _
                    OrElse rowARTCUST2.Item("CUST_SHIP_TO_ADDR1") & String.Empty <> CUST_SHIP_TO_ADDR1 _
                    OrElse rowARTCUST2.Item("CUST_SHIP_TO_ADDR2") & String.Empty <> CUST_SHIP_TO_ADDR2 _
                    OrElse rowARTCUST2.Item("CUST_SHIP_TO_ADDR3") & String.Empty <> CUST_SHIP_TO_ADDR3 _
                    OrElse rowARTCUST2.Item("CUST_SHIP_TO_CITY") & String.Empty <> CUST_SHIP_TO_CITY _
                    OrElse rowARTCUST2.Item("CUST_SHIP_TO_STATE") & String.Empty <> CUST_SHIP_TO_STATE _
                    OrElse rowARTCUST2.Item("CUST_SHIP_TO_ZIP_CODE") & String.Empty <> CUST_SHIP_TO_ZIP_CODE _
                    OrElse rowARTCUST2.Item("CUST_SHIP_TO_COUNTRY") & String.Empty <> CUST_SHIP_TO_COUNTRY Then

                    sql = "Update ARTCUST2 SET "
                    sql &= "   CUST_SHIP_TO_PHONE = '" & CUST_SHIP_TO_PHONE & "'"
                    sql &= " , CUST_SHIP_TO_URL = '" & CUST_SHIP_TO_URL & "'"
                    sql &= " , CUST_SHIP_TO_NAME = '" & CUST_SHIP_TO_NAME & "'"
                    sql &= " , CUST_SHIP_TO_ADDR1 = '" & CUST_SHIP_TO_ADDR1 & "'"
                    sql &= " , CUST_SHIP_TO_ADDR2 = '" & CUST_SHIP_TO_ADDR2 & "'"
                    sql &= " , CUST_SHIP_TO_ADDR3 = '" & CUST_SHIP_TO_ADDR3 & "'"
                    sql &= " , CUST_SHIP_TO_CITY = '" & CUST_SHIP_TO_CITY & "'"
                    sql &= " , CUST_SHIP_TO_STATE = '" & CUST_SHIP_TO_STATE & "'"
                    sql &= " , CUST_SHIP_TO_ZIP_CODE = '" & CUST_SHIP_TO_ZIP_CODE & "'"
                    sql &= " , CUST_SHIP_TO_COUNTRY = '" & CUST_SHIP_TO_COUNTRY & "'"
                    sql &= " Where CUST_CODE = :PARM1 AND CUST_SHIP_TO_NO = :PARM2"
                    ABSolution.ASCDATA1.ExecuteSQL(sql, "VV", New Object() {CUST_CODE, CUST_SHIP_TO_NO})
                End If

            Catch ex As Exception
                RecordLogEntry("UpdateCustomerShipTo: " & ex.Message)
            End Try

        End Function

        ''' <summary>
        ''' Updates the Freight for an Order
        ''' </summary>
        ''' <param name="ORDR_NO"></param>
        ''' <remarks></remarks>
        Private Function UpdateSalesOrderTotal(ByVal ORDR_NO As String) As Boolean

            Dim rowSOTORDR1 As DataRow = Nothing

            Try
                If testMode Then RecordLogEntry("Enter UpdateSalesOrderTotal. " & ORDR_NO)

                rowSOTORDR1 = dst.Tables("SOTORDR1").Rows(0)
                Dim tblSOTORDR2 As DataTable = dst.Tables("SOTORDR2")
                Dim ORDR_FREIGHT As Double = 0

                Dim ORDR_TOTAL_QTY As Integer = Val(dst.Tables("SOTORDR2").Compute("SUM(ORDR_QTY)", "ORDR_NO = '" & ORDR_NO & "'") & String.Empty)
                Dim ORDR_TOTAL_SALES As Double = Val(dst.Tables("SOTORDR2").Compute("SUM(ORDR_LNO_EXT)", "ORDR_NO = '" & ORDR_NO & "'") & String.Empty)
                Dim shipToPatient As Boolean = rowSOTORDR1.Item("ORDR_DPD") & String.Empty = "1"

                rowSOTORDR1.Item("ORDR_SALES") = Math.Round(ORDR_TOTAL_SALES, 2, MidpointRounding.AwayFromZero)

                If rowSOTORDR1.Item("ORDR_NO_FREIGHT") & String.Empty <> "1" Then
                    ORDR_FREIGHT = TAC.SOCMAIN1.Get_INV_FREIGHT(baseClass, rowSOTORDR1.Item("CUST_CODE"), _
                        rowSOTORDR1.Item("CUST_SHIP_TO_NO") & "", _
                        rowSOTORDR1.Item("SHIP_VIA_CODE"), rowSOTORDR1.Item("ORDR_DATE"), ORDR_TOTAL_QTY, ORDR_TOTAL_SALES, shipToPatient, "E")
                End If

                rowSOTORDR1.Item("ORDR_FREIGHT") = Math.Round(ORDR_FREIGHT, 2, MidpointRounding.AwayFromZero)

                rowSOTORDR1.Item("ORDR_STAX") = Me.GetOrderSalesTaxByState(rowSOTORDR1, dst.Tables("SOTORDR2"))
                If rowSOTORDR1.Item("ORDR_STAX") = -1 Then
                    rowSOTORDR1.Item("ORDR_STAX") = 0
                    Return False
                End If

                rowSOTORDR1.Item("ORDR_TOTAL_AMT") = Val(rowSOTORDR1.Item("ORDR_SALES") & String.Empty) _
                    + Val(rowSOTORDR1.Item("ORDR_FREIGHT") & String.Empty) _
                    + Val(rowSOTORDR1.Item("ORDR_STAX") & String.Empty) _
                    + Val(rowSOTORDR1.Item("ORDR_SAMPLE_SURCHARGE") & String.Empty) _
                    + Val(rowSOTORDR1.Item("ORDR_MISC_CHG_AMT") & String.Empty)

                If testMode Then RecordLogEntry("Exit UpdateSalesOrderTotal. " & ORDR_NO)
                Return True

            Catch ex As Exception
                RecordLogEntry("UpdateSalesOrderTotal: (" & ORDR_NO & ") " & ex.Message)
                Return False
            End Try
        End Function

        ''' <summary>
        ''' Looks for DPD order without an Annual Supply.
        ''' </summary>
        ''' <param name="rowSOTORDR1"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Private Function DpdOrderWithoutAnAnnualSupply(ByRef rowSOTORDR1 As DataRow, ByRef tblSOTORDR2 As DataTable) As String

            Try
                ' Must be DPD
                If rowSOTORDR1.Item("ORDR_DPD") & String.Empty <> "1" Then Return String.Empty

                ' Must be Vision Web
                If rowSOTORDR1.Item("ORDR_SOURCE") & String.Empty <> "V" Then Return String.Empty

                ' Must have Revenue
                If tblSOTORDR2.Select("ORDR_UNIT_PRICE > 0").Length = 0 Then Return String.Empty

                ' Must have ASP Code
                If tblSOTORDR2.Select("ASP_CODE IS NULL OR ASP_CODE = ''").Length = 0 Then Return String.Empty

                ' This is so it does not happen all the time, only initially for each processing cycle
                If aspPriceCatgys.Count = 0 Then
                    aspPriceCatgys = (From r In ABSolution.ASCDATA1.GetDataTable("SELECT DISTINCT PRICE_CATGY_CODE FROM PPTASPM1 P1 JOIN PPTASPM2 P2 USING (ASP_CODE) WHERE P1.ASP_STATUS='A' AND NVL(P1.ASP_DATE_END,TO_DATE('31-DEC-9999')) >= SYSDATE").AsEnumerable() _
                         Select (r.Item("PRICE_CATGY_CODE").ToString())).ToList()
                End If

                If rowWBTPARM1 Is Nothing Then
                    rowWBTPARM1 = ABSolution.ASCDATA1.GetDataRow("SELECT * FROM WBTPARM1 WHERE WB_PARM_KEY = :PARM1", "V", New Object() {"Z"})
                End If

                'make sure qty ordered > min asp qty
                Dim x = (From r In tblSOTORDR2.AsEnumerable() _
                                   Group By PG = r.Item("PATIENT_GROUP") Into OQ = Sum(Val(r.Item("ORDR_QTY"))) _
                                   Where OQ >= Val(rowWBTPARM1.Item("MIN_ASP_QTY"))).Count()
                If x > 0 Then
                    'make sure price catgys ordered exist in asp table
                    Dim orderPriceCatgys As List(Of String) = (From r In tblSOTORDR2.AsEnumerable() _
                                                             Select (r.Item("PRICE_CATGY_CODE").ToString()) Distinct).ToList()
                    If orderPriceCatgys.Intersect(aspPriceCatgys).Count > 0 Then
                        Return DPDNoAnnualSupply
                    End If
                End If

                Return String.Empty
            Catch ex As Exception
                RecordLogEntry("DpdOrderWithoutAnAnnualSupply: " & ex.Message)
                Return String.Empty
            End Try

            Return String.Empty
        End Function

        Private Function DpdOrderCODneedsAuthorization(ByRef rowSOTORDR1 As DataRow, ByRef tblSOTORDR2 As DataTable) As String

            Try
                ' Must be DPD
                If rowSOTORDR1.Item("ORDR_DPD") & String.Empty <> "1" Then Return String.Empty

                ' Must be Vision Web
                If rowSOTORDR1.Item("ORDR_SOURCE") & String.Empty <> "V" Then Return String.Empty

                ' Must have Revenue
                If tblSOTORDR2.Select("ORDR_UNIT_PRICE > 0").Length = 0 Then Return String.Empty

                Dim rowTATTERM1 As DataRow = baseClass.LookUp("TATTERM1", rowSOTORDR1.Item("TERM_CODE") & String.Empty)
                If rowTATTERM1 Is Nothing Then Return String.Empty
                If rowTATTERM1.Item("TERM_TYPE") = "A" Then Return String.Empty

                Return DpdCODCustomer

            Catch ex As Exception
                RecordLogEntry("DpdOrderCODneedsAuthorization: " & ex.Message)
                Return String.Empty
            End Try
        End Function

        Private Function CreateDPDPatientRecord(ByVal CUST_CODE As String, ByVal CUST_SHIP_TO_NO As String, ByVal rowSOTORDR5 As DataRow) As String

            Dim Sql As String = String.Empty
            Dim rowARTCUST4 As DataRow = Nothing
            Dim patientNo As String

            If CUST_SHIP_TO_NO.StartsWith("x") Then
                CUST_SHIP_TO_NO = String.Empty
            End If

            Try
                '' Determine if the DPD Patient exists for this customer
                'Sql = "SELECT * FROM ARTCUST4  " & _
                '      " WHERE PATIENT_NAME = :PARM1" & _
                '      " AND PATIENT_ADDR1 = :PARM2" & _
                '      " AND DECODE(PATIENT_ADDR2,NVL(:PARM3,PATIENT_ADDR2),'1','0') = '1'" & _
                '      " AND DECODE(PATIENT_CITY,NVL(:PARM4,PATIENT_CITY),'1','0') = '1'" & _
                '      " AND DECODE(PATIENT_STATE,NVL(:PARM5,PATIENT_STATE),'1','0') = '1'" & _
                '      " AND DECODE(PATIENT_ZIP_CODE,NVL(:PARM6,PATIENT_ZIP_CODE),'1','0') = '1'" & _
                '      " AND CUST_CODE = :PARM7" & _
                '      " AND DECODE(CUST_SHIP_TO_NO,NVL(:PARM8,CUST_SHIP_TO_NO),'1','0') = '1'"

                'rowARTCUST4 = ABSolution.ASCDATA1.GetDataRow(Sql, "VVVVVVVV", New Object() { _
                '                rowSOTORDR5.Item("CUST_NAME") & String.Empty, _
                '                rowSOTORDR5.Item("CUST_ADDR1") & String.Empty, _
                '                rowSOTORDR5.Item("CUST_ADDR2") & String.Empty, _
                '                rowSOTORDR5.Item("CUST_CITY") & String.Empty, _
                '                rowSOTORDR5.Item("CUST_STATE") & String.Empty, _
                '                rowSOTORDR5.Item("CUST_ZIP_CODE") & String.Empty, _
                '                CUST_CODE, CUST_SHIP_TO_NO})

                Sql = "SELECT * FROM ARTCUST4  " & _
                      " WHERE UPPER(PATIENT_NAME) = :PARM1" & _
                      " AND UPPER(PATIENT_ADDR1) = :PARM2" & _
                      " AND DECODE(UPPER(PATIENT_ADDR2), NVL(:PARM3, UPPER(PATIENT_ADDR2)), '1', '0') = '1'" & _
                      " AND DECODE(PATIENT_ZIP_CODE,NVL(:PARM4,PATIENT_ZIP_CODE),'1','0') = '1'" & _
                      " AND CUST_CODE = :PARM5" & _
                      " AND DECODE(CUST_SHIP_TO_NO,NVL(:PARM6,CUST_SHIP_TO_NO),'1','0') = '1'"

                rowARTCUST4 = ABSolution.ASCDATA1.GetDataRow(Sql, "VVVVVV", New Object() { _
                                (rowSOTORDR5.Item("CUST_NAME") & String.Empty).ToString.ToUpper, _
                                (rowSOTORDR5.Item("CUST_ADDR1") & String.Empty).ToString.ToUpper, _
                                (rowSOTORDR5.Item("CUST_ADDR2") & String.Empty).ToString.ToUpper, _
                                rowSOTORDR5.Item("CUST_ZIP_CODE") & String.Empty, _
                                CUST_CODE, CUST_SHIP_TO_NO})


                If rowARTCUST4 Is Nothing Then
                    patientNo = ABSolution.ASCMAIN1.Next_Control_No("ARTCUST4.PATIENT_NO")

                    rowARTCUST4 = dst.Tables("ARTCUST4").NewRow
                    rowARTCUST4.Item("PATIENT_NO") = patientNo
                    rowARTCUST4.Item("PATIENT_NAME") = rowSOTORDR5.Item("CUST_NAME") & String.Empty
                    rowARTCUST4.Item("PATIENT_ADDR1") = rowSOTORDR5.Item("CUST_ADDR1") & String.Empty
                    rowARTCUST4.Item("PATIENT_ADDR2") = rowSOTORDR5.Item("CUST_ADDR2") & String.Empty
                    rowARTCUST4.Item("PATIENT_CITY") = rowSOTORDR5.Item("CUST_CITY") & String.Empty
                    rowARTCUST4.Item("PATIENT_STATE") = rowSOTORDR5.Item("CUST_STATE") & String.Empty
                    rowARTCUST4.Item("PATIENT_ZIP_CODE") = rowSOTORDR5.Item("CUST_ZIP_CODE") & String.Empty
                    rowARTCUST4.Item("INIT_OPER") = ABSolution.ASCMAIN1.USER_ID
                    rowARTCUST4.Item("INIT_DATE") = DateTime.Now
                    rowARTCUST4.Item("PATIENT_STATUS") = "A"
                    rowARTCUST4.Item("CUST_CODE") = CUST_CODE

                    If CUST_SHIP_TO_NO.Length > 0 Then
                        rowARTCUST4.Item("CUST_SHIP_TO_NO") = CUST_SHIP_TO_NO
                    End If

                    dst.Tables("ARTCUST4").Rows.Add(rowARTCUST4)
                  Else
                    patientNo = rowARTCUST4.Item("PATIENT_NO") & String.Empty
                End If

                Return patientNo

            Catch ex As Exception
                Return String.Empty
            End Try

        End Function

        Private Function ValidateDPDAddress(ByRef rowSOTORDR1 As DataRow, ByRef rowSOTORDR5_ST As DataRow) As Boolean

            ValidateDPDAddress = False

            Try
                RecordLogEntry("Enter ValidateDPDAddress")
                Dim clsSHCUPSC1 As New TAC.SHCUPSC1
                Dim addressValidations As New List(Of TAC.SHCUPSC1.AddressValidationResponse)
                Dim errMsg As String = String.Empty

                Dim AddressLine1 As String = (rowSOTORDR5_ST.Item("CUST_ADDR1") & String.Empty).ToString.Trim.ToUpper
                Dim AddressLine2 As String = (rowSOTORDR5_ST.Item("CUST_ADDR2") & String.Empty).ToString.Trim.ToUpper
                Dim AddressLine3 As String = (rowSOTORDR5_ST.Item("CUST_ADDR3") & String.Empty).ToString.Trim.ToUpper
                Dim Name As String = (rowSOTORDR5_ST.Item("CUST_NAME") & String.Empty).ToString.Trim.ToUpper
                Dim CompanyName As String = (rowSOTORDR5_ST.Item("CUST_NAME") & String.Empty).ToString.Trim.ToUpper

                Dim City As String = (rowSOTORDR5_ST.Item("CUST_CITY") & String.Empty).ToString.Trim.ToUpper
                Dim State As String = (rowSOTORDR5_ST.Item("CUST_STATE") & String.Empty).ToString.Trim.ToUpper
                Dim PostalCode As String = (rowSOTORDR5_ST.Item("CUST_ZIP_CODE") & String.Empty).ToString.Trim.ToUpper

                If City.Length = 0 OrElse State.Length = 0 OrElse PostalCode.Length = 0 Then
                    Return False
                End If

                If dst.Tables("SOTORDR1").Rows(0).Item("ORDR_DPD") & String.Empty = "1" Then
                    CompanyName = String.Empty
                Else
                    Name = rowSOTORDR5_ST.Item("CUST_CONTACT") & String.Empty
                End If

                If PostalCode.Length > 5 Then
                    PostalCode = PostalCode.Substring(0, 5)
                End If

                ' Needed since service cannot access S drive
                Dim subDirectory As String = DateTime.Now.ToString("yyyyMMdd") & "\"
                Dim rowSOTPARM2 As DataRow = ABSolution.ASCDATA1.GetDataRow("Select * From SOTPARM2 Where SO_PARM_KEY = 'Z'", False)

                Dim overrideCarrierRequestXmlDir As String = String.Empty
                Dim overrideCarrierResponseXmlDir As String = String.Empty
                Dim overrideCarrierResponseRptDir As String = String.Empty

                If rowSOTPARM2 IsNot Nothing Then
                    overrideCarrierRequestXmlDir = rowSOTPARM2.Item("SO_PARM_REQ_XML_DIR") & String.Empty
                    overrideCarrierResponseXmlDir = rowSOTPARM2.Item("SO_PARM_RESP_XML_DIR") & String.Empty
                    overrideCarrierResponseRptDir = rowSOTPARM2.Item("SO_PARM_RESP_RPT_DIR") & String.Empty

                    overrideCarrierRequestXmlDir = overrideCarrierRequestXmlDir.Trim.ToUpper
                    If Not overrideCarrierRequestXmlDir.EndsWith("\") Then overrideCarrierRequestXmlDir &= "\"

                    overrideCarrierResponseXmlDir = overrideCarrierResponseXmlDir.Trim.ToUpper
                    If Not overrideCarrierResponseXmlDir.EndsWith("\") Then overrideCarrierResponseXmlDir &= "\"

                    overrideCarrierResponseRptDir = overrideCarrierResponseRptDir.Trim.ToUpper
                    If Not overrideCarrierResponseRptDir.EndsWith("\") Then overrideCarrierResponseRptDir &= "\"

                    overrideCarrierRequestXmlDir &= subDirectory
                    overrideCarrierResponseXmlDir &= subDirectory
                    overrideCarrierResponseRptDir &= subDirectory
                End If

                For Each Path As String In New String() {overrideCarrierRequestXmlDir, overrideCarrierResponseXmlDir, overrideCarrierResponseRptDir}
                    Try
                        If Not My.Computer.FileSystem.DirectoryExists(Path) Then
                            My.Computer.FileSystem.CreateDirectory(Path)
                        End If
                    Catch ex As Exception

                    End Try
                Next

                If convert AndAlso overrideCarrierRequestXmlDir.StartsWith(DriveLetter) Then
                    clsSHCUPSC1.CarrierRequestXmlDir = overrideCarrierRequestXmlDir
                    'RecordLogEntry("convert clsSHCUPSC1.CarrierRequestXmlDir: ")
                    clsSHCUPSC1.CarrierRequestXmlDir = clsSHCUPSC1.CarrierRequestXmlDir.Replace(DriveLetter, DriveLetterIP)
                End If

                If convert AndAlso overrideCarrierResponseRptDir.StartsWith(DriveLetter) Then
                    clsSHCUPSC1.CarrierResponseRptDir = overrideCarrierResponseRptDir
                    'RecordLogEntry("convert clsSHCUPSC1.CarrierResponseRptDir: ")
                    clsSHCUPSC1.CarrierResponseRptDir = clsSHCUPSC1.CarrierResponseRptDir.Replace(DriveLetter, DriveLetterIP)
                End If

                If convert AndAlso overrideCarrierResponseXmlDir.StartsWith(DriveLetter) Then
                    clsSHCUPSC1.CarrierResponseXmlDir = overrideCarrierResponseXmlDir
                    'RecordLogEntry("convert clsSHCUPSC1.CarrierResponseXmlDir: ")
                    clsSHCUPSC1.CarrierResponseXmlDir = clsSHCUPSC1.CarrierResponseXmlDir.Replace(DriveLetter, DriveLetterIP)
                End If

                'RecordLogEntry("clsSHCUPSC1.CarrierRequestXmlDir: " & clsSHCUPSC1.CarrierRequestXmlDir)
                'RecordLogEntry("clsSHCUPSC1.CarrierResponseRptDir: " & clsSHCUPSC1.CarrierResponseRptDir)
                'RecordLogEntry("clsSHCUPSC1.CarrierResponseXmlDir: " & clsSHCUPSC1.CarrierResponseXmlDir)

                Dim validAddress As Boolean = clsSHCUPSC1.AddressVaildationRequest(AddressLine1, _
                                                                                   AddressLine2, _
                                                                                   AddressLine3, _
                                                                                   Name, _
                                                                                   City, _
                                                                                   State, _
                                                                                   PostalCode, _
                                                                                   CompanyName, _
                                                                                   addressValidations, errMsg)

                If errMsg.Length > 0 Then
                    RecordLogEntry("ValidateDPDAddress clsSHCUPSC1 Error: " & errMsg)
                    Return False
                ElseIf addressValidations.Count = 0 Then
                    RecordLogEntry("ValidateDPDAddress: Address Count = 0")
                    Return False
                ElseIf addressValidations.Count > 0 Then
                    ' If there is a match to what we sent then do not display the matches
                    For Each addressSel As TAC.SHCUPSC1.AddressValidationResponse In addressValidations
                        If addressSel.City.ToUpper = City.ToUpper AndAlso _
                                addressSel.State = State.ToUpper AndAlso _
                                addressSel.PostalCode = PostalCode.ToUpper Then

                            rowSOTORDR5_ST.Item("CUST_CITY") = addressSel.City
                            rowSOTORDR5_ST.Item("CUST_STATE") = addressSel.State
                            rowSOTORDR5_ST.Item("CUST_ZIP_CODE") = addressSel.PostalCode
                            RecordLogEntry("ValidateDPDAddress: Valid DPD Address")
                            Return True
                        End If
                    Next
                    RecordLogEntry("ValidateDPDAddress: Address not matched")
                End If

            Catch ex As Exception
                RecordLogEntry("ValidateDPDAddress :" & ex.Message)
                Return False
            Finally
                RecordLogEntry("Exit ValidateDPDAddress")
            End Try

        End Function

#End Region

#Region "DataSet Functions"

        Private Function ClearDataSetTables(ByVal ClearXSTtables As Boolean) As Boolean

            Try

                If testMode Then RecordLogEntry("Enter ClearDataSetTables.")

                With dst
                    .Tables("SOTORDRS").Clear()
                    .Tables("SOTORDR1").Clear()
                    .Tables("SOTORDR2").Clear()
                    .Tables("SOTORDR3").Clear()
                    .Tables("SOTORDR4").Clear()
                    .Tables("SOTORDR5").Clear()

                    .Tables("SOTORDRE").Clear()
                    .Tables("SOTORDRP").Clear()
                    .Tables("SOTORDRW").Clear()
                    .Tables("SOTORDRX").Clear()

                    .Tables("XSTORDRQ").Clear()
                    .Tables("ARTCUST4").Clear()

                    .Tables("DETJOBM1").Clear()
                    .Tables("DETJOBM2").Clear()
                    .Tables("DETJOBM3").Clear()
                    .Tables("DETJOBIE").Clear()

                    If ClearXSTtables Then
                        .Tables("XSTORDR1").Clear()
                        .Tables("XSTORDR2").Clear()

                        .Tables("XMTORDR1").Clear()
                        .Tables("XMTORDR2").Clear()
                    End If


                End With

                rowARTCUST1 = Nothing
                rowARTCUST2 = Nothing
                rowARTCUST3 = Nothing
                rowICTITEM1 = Nothing
                rowSOTSVIA1 = Nothing
                rowTATTERM1 = Nothing

                If testMode Then RecordLogEntry("Exit ClearDataSetTables.")
                Return True

            Catch ex As Exception
                RecordLogEntry("ClearDataSetTables: " & ex.Message)
                Return False
            End Try

        End Function

        Private Sub Dependent_Updates(ByVal ORDR_NO As String, ByVal S As Integer)

            Dim PLUS_OR_MINUS As String = "+1*"
            Dim sql As String = String.Empty

            If S = -1 Then
                PLUS_OR_MINUS = "-1*"
            End If

            If testMode Then RecordLogEntry("Enter Dependent_Updates.")

            sql = "" _
            & "BEGIN DECLARE CURSOR C1 IS " _
            & " SELECT SOTORDR2.ITEM_CODE,SOTORDR1.WHSE_CODE,SOTORDR2.ORDR_QTY_OPEN" _
            & " FROM SOTORDR2,SOTORDR1 WHERE SOTORDR2.ORDR_NO = SOTORDR1.ORDR_NO" _
            & " AND SOTORDR2.ORDR_NO = '" & ORDR_NO & "';" _
            & " BEGIN FOR R1 IN C1 LOOP" _
            & " UPDATE ICTSTAT2 SET WHSE_QTY_OPEN = NVL(WHSE_QTY_OPEN,0) " & PLUS_OR_MINUS & " NVL(R1.ORDR_QTY_OPEN,0)" _
            & " WHERE ITEM_CODE = R1.ITEM_CODE AND WHSE_CODE = R1.WHSE_CODE;" _
            & " IF SQL%NOTFOUND THEN" _
            & " INSERT INTO ICTSTAT2 (ITEM_CODE, WHSE_CODE, WHSE_QTY_OPEN)" _
            & " VALUES (R1.ITEM_CODE,R1.WHSE_CODE," & PLUS_OR_MINUS & " NVL(R1.ORDR_QTY_OPEN,0));" _
            & " END IF; " _
            & " END LOOP; END; END;"
            ABSolution.ASCDATA1.ExecuteSQL(sql)

            If S = 1 Then
                sql = "" _
                & " BEGIN DECLARE CURSOR C1 IS" _
                & " SELECT SOTORDRP.PROM_CODE_INITIAL_ORDER PROM_CODE" _
                & " , SOTORDR1.CUST_CODE" _
                & " , SYSDATE + PPTPROM1.PROM_PROTECT_DAYS PRICE_PROTECTED_UNTIL" _
                & " , SOTORDRP.ORDR_NO ORDR_NO_INITIAL" _
                & " , SOTORDRP.STKB_LEVEL" _
                & " FROM SOTORDRP,PPTPROM1,SOTORDR1" _
                & " WHERE PPTPROM1.PROM_CODE = SOTORDRP.PROM_CODE_INITIAL_ORDER" _
                & " AND SOTORDR1.ORDR_NO = SOTORDRP.ORDR_NO" _
                & " AND SOTORDRP.PROM_CODE_INITIAL_ORDER IS NOT NULL" _
                & " AND SOTORDRP.ORDR_NO = '" & ORDR_NO & "';" _
                & " BEGIN FOR R1 IN C1 LOOP" _
                & " DELETE FROM PPTPROM3 WHERE PROM_CODE = R1.PROM_CODE AND CUST_CODE = R1.CUST_CODE;" _
                & " INSERT INTO PPTPROM3 " _
                & " (PROM_CODE, CUST_CODE, PRICE_PROTECTED_UNTIL, ORDR_NO_INITIAL, STKB_LEVEL)" _
                & " VALUES " _
                & " (R1.PROM_CODE, R1.CUST_CODE, R1.PRICE_PROTECTED_UNTIL, R1.ORDR_NO_INITIAL, R1.STKB_LEVEL);" _
                & " END LOOP; END; END;"
                ABSolution.ASCDATA1.ExecuteSQL(sql)
            Else
                sql = "Delete from PPTPROM3 where ORDR_NO_INITIAL = '" & ORDR_NO & "'"
                ABSolution.ASCDATA1.ExecuteSQL(sql)
            End If

            If testMode Then RecordLogEntry("Exit Dependent_Updates.")

        End Sub

        Private Sub LoadTablesForPricing()

            With dst

                If testMode Then RecordLogEntry("Enter LoadTablesForPricing.")

                .Tables("SOTORDR1").Columns.Add("ORDR_REL_HOLD_CODES", GetType(System.String))

                .Tables("SOTORDR2").Columns.Add("QTY_ONH", GetType(System.Int32))
                .Tables("SOTORDR2").Columns.Add("QTY_AVA", GetType(System.Int32))
                .Tables("SOTORDR2").Columns.Add("QTY_OPO", GetType(System.Int32))
                .Tables("SOTORDR2").Columns.Add("LINE_AMOUNT", GetType(System.Decimal), "ISNULL(ORDR_QTY,0) * ISNULL(ORDR_UNIT_PRICE,0)")
                .Tables("SOTORDR2").Columns.Add("PO_ORDER_NO")
                .Tables("SOTORDR2").Columns.Add("PO_ORDER_LNO", GetType(System.Int32))
                .Tables("SOTORDR2").Columns.Add("FBO_DATE_EXPECTED", GetType(System.DateTime))
                .Tables("SOTORDR2").Columns.Add("PO_QTY_OPN", GetType(System.Int32))
                .Tables("SOTORDR2").Columns.Add("ASP_MESSAGE")
                .Tables("SOTORDR2").Columns.Add("ORDR_REL_HOLD_CODES", GetType(System.String))
                .Tables("SOTORDR2").Columns.Add("SAMPLE_IND", GetType(System.String))
                .Tables("SOTORDR2").Columns.Add("PRICE_CATGY_SAMPLE_IND", GetType(System.String))

                clsSOCORDR1 = New TAC.SOCORDR1(SOTINVH2_PC, SOTORDRP, SOTORDR2_pricing, baseClass.clsASCBASE1)

                If testMode Then RecordLogEntry("Exit LoadTablesForPricing.")

            End With

        End Sub

        Private Function PrepareDatasetEntries() As Boolean

            Try

                Dim sql As String = String.Empty
                If testMode Then RecordLogEntry("Enter PrepareDatasetEntries.")

                dst = baseClass.clsASCBASE1.dst
                dst.Tables.Clear()

                With dst

                    baseClass.Create_TDA(.Tables.Add, "SOTORDRS", "*", 1)
                    baseClass.Create_TDA(.Tables.Add, "SOTORDR1", "*")
                    baseClass.Create_TDA(.Tables.Add, "SOTORDR2", "*", 1)

                    baseClass.Create_TDA(.Tables.Add, "SOTORDRX", "*", 2)
                    baseClass.Create_TDA(.Tables.Add, "SOTSVIAF", "*", 1)

                    baseClass.Create_TDA(.Tables.Add, "DETJOBM1", "*")
                    baseClass.Create_TDA(.Tables.Add, "DETJOBM2", "*")
                    .Tables("DETJOBM2").Columns.Add("JOB_AMT", GetType(System.Double), "JOB_QTY * JOB_PRICE")
                    .Tables("DETJOBM2").Columns.Add("CUST_PRICE", GetType(System.Double))

                    baseClass.Create_TDA(.Tables.Add, "DETJOBM3", "*")
                    baseClass.Create_TDA(.Tables.Add, "DETJOBM4", "*")
                    baseClass.Create_TDA(.Tables.Add, "DETJOBIE", "*")

                    With .Tables("SOTORDR2")
                        .Columns.Add("ORDR_LNO_EXT", GetType(System.Double), "ISNULL(ORDR_QTY, 0) * ISNULL(ORDR_UNIT_PRICE, 0)")
                    End With

                    baseClass.Create_TDA(.Tables.Add, "SOTORDR3", "*", 1)
                    baseClass.Create_TDA(.Tables.Add, "SOTORDR5", "*", 1)

                    baseClass.Create_TDA(.Tables.Add, "SOTORDRW", "*", 1)
                    baseClass.Create_TDA(.Tables.Add, "SOTORDRE", "*", 0)

                    ' Web Service Import
                    baseClass.Create_TDA(.Tables.Add, "XSTORDR1", "*", 1)
                    baseClass.Create_TDA(.Tables.Add, "XSTORDR2", "*", 1)

                    baseClass.Create_TDA(.Tables.Add, "XMTXREF1", "*")
                    baseClass.clsASCBASE1.Fill_Records("XMTXREF1", String.Empty, True, "Select * From XMTXREF1")

                    baseClass.Create_TDA(.Tables.Add, "XMTORDR1", "*", 1)
                    baseClass.Create_TDA(.Tables.Add, "XMTORDR2", "*", 1)

                    baseClass.Create_TDA(.Tables.Add, "XSTORDRQ", "*", 2)
                    baseClass.Create_TDA(.Tables.Add, "ARTCUST4", "*")

                    baseClass.Create_TDA(dst.Tables.Add, "SOTORDRO", "Select LPAD( ' ', 15) ORDR_REL_HOLD_CODES, ORDR_COMMENT From SOTORDR1", , False)
                    CreateOrderRelCodesTable()

                    ' EDI Import 
                    baseClass.Create_TDA(.Tables.Add, "EDT850I0", "*", 1)
                    baseClass.Create_TDA(.Tables.Add, "EDT850I1", "*", 1)
                    baseClass.Create_TDA(.Tables.Add, "EDT850I2", "*", 1)
                    baseClass.Create_TDA(.Tables.Add, "EDT850I5", "*", 1)
                    baseClass.Create_TDA(.Tables.Add, "EDT850I1D", "Select EDI_ISA_NO From EDT850I1", , False)

                    LoadTablesForPricing()

                    'baseClass.Get_PARM("SOTPARM1")
                    baseClass.Create_TDA(.Tables.Add, "SOTPARM1", "*")
                    baseClass.clsASCBASE1.Fill_Records("SOTPARM1", "Z")

                    baseClass.Create_TDA(.Tables.Add, "SOTPARMB", "*")
                    baseClass.clsASCBASE1.Fill_Records("SOTPARMB", "Z")

                    baseClass.Create_TDA(.Tables.Add, "DETPARM1", "*")
                    baseClass.clsASCBASE1.Fill_Records("DETPARM1", "Z")

                    baseClass.Create_TDA(.Tables.Add, "DETCOLMC", "*")
                    baseClass.clsASCBASE1.Fill_Records("DETCOLMC", "", True, "SELECT * FROM DETCOLMC")

                    SO_PARM_SHIP_ND = (dst.Tables("SOTPARMB").Rows(0).Item("SO_PARM_SHIP_ND") & String.Empty).ToString.Trim
                    SO_PARM_SHIP_ND_DPD = (dst.Tables("SOTPARMB").Rows(0).Item("SO_PARM_SHIP_ND_DPD") & String.Empty).ToString.Trim
                    SO_PARM_SHIP_ND_COD = (dst.Tables("SOTPARMB").Rows(0).Item("SO_PARM_SHIP_ND_COD") & String.Empty).ToString.Trim

                    If dst.Tables("DETPARM1") IsNot Nothing AndAlso dst.Tables("DETPARM1").Rows.Count > 0 Then
                        DE_PARM_FOG_FREE_COST = Val(dst.Tables("DETPARM1").Rows(0).Item("DE_PARM_FOG_FREE_COST") & String.Empty)
                    Else
                        DE_PARM_FOG_FREE_COST = 45
                    End If

                    If dst.Tables("SOTPARM1") IsNot Nothing AndAlso dst.Tables("SOTPARM1").Rows.Count > 0 Then
                        DpdDefaultShipViaCode = (dst.Tables("SOTPARM1").Rows(0).Item("SO_PARM_SHIP_VIA_CODE_DPD") & String.Empty).ToString.Trim
                    End If

                    sql = "ITEM_CODE = :PARM1 OR ITEM_UPC_CODE = :PARM1 OR ITEM_OPC_CODE = :PARM1 OR ITEM_EAN_CODE = :PARM1"
                    baseClass.Create_Lookup("ICTITEM1", "*", sql, "V", False)

                    sql = "ITEM_CODE = :PARM1 OR ITEM_PROD_ID = :PARM1"
                    baseClass.Create_Lookup("ICTCATL1", "*", sql, "V", False)

                    sql = "CUST_CODE = :PARM1"
                    baseClass.Create_Lookup("ARTCUST1", "*", sql, "V", False)

                    sql = "CUST_CODE = :PARM1 AND CUST_SHIP_TO_NO = :PARM2"
                    baseClass.Create_Lookup("ARTCUST2", "*", sql, "VV", False)

                    sql = "CUST_CODE = :PARM1 AND NVL(FRT_CONT_DATE_START, SYSDATE) <= SYSDATE AND NVL(FRT_CONT_DATE_END, SYSDATE) >= SYSDATE"
                    sql &= " AND (CUST_CODE, FRT_CONT_NO) IN (SELECT CUST_CODE, MAX(FRT_CONT_NO) FROM ARTCUST3 GROUP BY CUST_CODE)"
                    baseClass.Create_Lookup("ARTCUST3", "*", sql, "V", False)

                    sql = "STAX_CODE = :PARM1"
                    baseClass.Create_Lookup("ARTSTAX1", "*", sql, "V", False)

                    sql = "SHIP_VIA_CODE = :PARM1"
                    baseClass.Create_Lookup("SOTSVIA1", "*", sql, "V", False)

                    sql = "TERM_CODE = :PARM1"
                    baseClass.Create_Lookup("TATTERM1", "*", sql, "V", False)

                    baseClass.Create_Lookup("SOTSVIA1")

                    STAX_CODE_states = New List(Of String)
                    sql = "Select Distinct STATE from ARTSTAX2"
                    For Each row As DataRow In ABSolution.ASCDATA1.GetDataTable(sql).Rows
                        STAX_CODE_states.Add(row.Item(0))
                    Next

                End With

                If testMode Then RecordLogEntry("Exit PrepareDatasetEntries.")
                Return True

            Catch ex As Exception
                RecordLogEntry("PrepareDatasetEntries: " & ex.Message)
                Return False
            End Try

        End Function

        Private Sub UpdateDataSetTables()

            Dim sql As String = String.Empty

            With baseClass
                Try
                    If testMode Then RecordLogEntry("Enter UpdateDataSetTables.")

                    .BeginTrans()
                    .clsASCBASE1.Update_Record_TDA("SOTORDR1")
                    .clsASCBASE1.Update_Record_TDA("SOTORDR2")
                    .clsASCBASE1.Update_Record_TDA("SOTORDR3")
                    .clsASCBASE1.Update_Record_TDA("SOTORDR5")

                    .clsASCBASE1.Update_Record_TDA("SOTORDRE")
                    .clsASCBASE1.Update_Record_TDA("ARTCUST4")

                    dst.Tables("SOTORDR4").AcceptChanges()
                    For Each row As DataRow In dst.Tables("SOTORDR4").Rows
                        row.SetAdded()
                    Next
                    .clsASCBASE1.Update_Record_TDA("SOTORDR4")

                    dst.Tables("SOTORDRP").AcceptChanges()
                    For Each row As DataRow In dst.Tables("SOTORDRP").Rows
                        row.SetAdded()
                    Next
                    .clsASCBASE1.Update_Record_TDA("SOTORDRP")

                    ' Log Errors on the order
                    dst.Tables("SOTORDRW").Clear()

                    Dim rowSOTORDR1 As DataRow = dst.Tables("SOTORDR1").Rows(0)
                    For Each holdCode As Char In (rowSOTORDR1.Item("ORDR_REL_HOLD_CODES") & String.Empty).ToString.Trim
                        Me.CreateOrderErrorRecord(rowSOTORDR1.Item("ORDR_NO"), 0, holdCode)
                    Next

                    For Each rowSOTORDR2 As DataRow In dst.Tables("SOTORDR2").Rows
                        For Each holdCode As Char In (rowSOTORDR2.Item("ORDR_REL_HOLD_CODES") & String.Empty).ToString.Trim
                            Me.CreateOrderErrorRecord(rowSOTORDR2.Item("ORDR_NO"), rowSOTORDR2.Item("ORDR_LNO"), holdCode)
                        Next
                    Next

                    .clsASCBASE1.Update_Record_TDA("SOTORDRW")

                    If rowSOTORDR1.Item("ORDR_STATUS") & String.Empty = "O" AndAlso rowSOTORDR1.Item("ORDR_REL_HOLD_CODES") & String.Empty = String.Empty Then
                        Dependent_Updates(rowSOTORDR1.Item("ORDR_NO"), 1)
                        ABSolution.ASCDATA1.ExecuteSQL("Begin ARPCUST6_ORDR_NO('ODG', '" & rowSOTORDR1.Item("CUST_CODE") & "', '" & rowSOTORDR1.Item("ORDR_NO") & "', SYSDATE, " & Val(rowSOTORDR1.Item("ORDR_SALES") & String.Empty) & ");  End;")
                    End If

                    .clsASCBASE1.Update_Record_TDA("XSTORDR1")
                    .clsASCBASE1.Update_Record_TDA("XSTORDR2")

                    .clsASCBASE1.Update_Record_TDA("XMTORDR1")
                    .clsASCBASE1.Update_Record_TDA("XMTORDR2")

                    .CommitTrans()
                    If testMode Then RecordLogEntry("Exit UpdateDataSetTables.")

                Catch ex As Exception
                    .Rollback()
                    RecordLogEntry("UpdateDataSetTables  : " & ex.Message)
                End Try
            End With

        End Sub

        ''' <summary>
        ''' Creates a Table of Order Release Hold Codes
        ''' </summary>
        ''' <remarks></remarks>
        Private Sub CreateOrderRelCodesTable()

            Dim rowSOTORDRO As DataRow = Nothing
            dst.Tables("SOTORDRO").Rows.Clear()

            ' Order Header Errors
            rowSOTORDRO = dst.Tables("SOTORDRO").NewRow
            rowSOTORDRO.Item("ORDR_REL_HOLD_CODES") = InvalidSoldTo
            rowSOTORDRO.Item("ORDR_COMMENT") = "Invalid Sold To"
            dst.Tables("SOTORDRO").Rows.Add(rowSOTORDRO)

            rowSOTORDRO = dst.Tables("SOTORDRO").NewRow
            rowSOTORDRO.Item("ORDR_REL_HOLD_CODES") = InvalidShipTo
            rowSOTORDRO.Item("ORDR_COMMENT") = "Invalid Ship To"
            dst.Tables("SOTORDRO").Rows.Add(rowSOTORDRO)

            rowSOTORDRO = dst.Tables("SOTORDRO").NewRow
            rowSOTORDRO.Item("ORDR_REL_HOLD_CODES") = InvalidShipVia
            rowSOTORDRO.Item("ORDR_COMMENT") = "Invalid Ship Via"
            dst.Tables("SOTORDRO").Rows.Add(rowSOTORDRO)

            rowSOTORDRO = dst.Tables("SOTORDRO").NewRow
            rowSOTORDRO.Item("ORDR_REL_HOLD_CODES") = ReviewSpecInstr
            rowSOTORDRO.Item("ORDR_COMMENT") = "Review Spec Instr"
            dst.Tables("SOTORDRO").Rows.Add(rowSOTORDRO)

            rowSOTORDRO = dst.Tables("SOTORDRO").NewRow
            rowSOTORDRO.Item("ORDR_REL_HOLD_CODES") = InvalidTaxCode
            rowSOTORDRO.Item("ORDR_COMMENT") = "Invalid Tax Code"
            dst.Tables("SOTORDRO").Rows.Add(rowSOTORDRO)

            rowSOTORDRO = dst.Tables("SOTORDRO").NewRow
            rowSOTORDRO.Item("ORDR_REL_HOLD_CODES") = InvalidTermsCode
            rowSOTORDRO.Item("ORDR_COMMENT") = "Invalid Terms Code"
            dst.Tables("SOTORDRO").Rows.Add(rowSOTORDRO)

            rowSOTORDRO = dst.Tables("SOTORDRO").NewRow
            rowSOTORDRO.Item("ORDR_REL_HOLD_CODES") = InvalidSalesTax
            rowSOTORDRO.Item("ORDR_COMMENT") = "Invalid Sales Tax"
            dst.Tables("SOTORDRO").Rows.Add(rowSOTORDRO)

            rowSOTORDRO = dst.Tables("SOTORDRO").NewRow
            rowSOTORDRO.Item("ORDR_REL_HOLD_CODES") = InvalidDPD
            rowSOTORDRO.Item("ORDR_COMMENT") = "Invalid DPD"
            dst.Tables("SOTORDRO").Rows.Add(rowSOTORDRO)

            rowSOTORDRO = dst.Tables("SOTORDRO").NewRow
            rowSOTORDRO.Item("ORDR_REL_HOLD_CODES") = InvalidPricing
            rowSOTORDRO.Item("ORDR_COMMENT") = "Invalid Pricing"
            dst.Tables("SOTORDRO").Rows.Add(rowSOTORDRO)

            rowSOTORDRO = dst.Tables("SOTORDRO").NewRow
            rowSOTORDRO.Item("ORDR_REL_HOLD_CODES") = InvalidSalesOrderTotal
            rowSOTORDRO.Item("ORDR_COMMENT") = "Invalid Sales Order Total"
            dst.Tables("SOTORDRO").Rows.Add(rowSOTORDRO)

            ' Order Detail Errors
            rowSOTORDRO = dst.Tables("SOTORDRO").NewRow
            rowSOTORDRO.Item("ORDR_REL_HOLD_CODES") = QtyOrdered
            rowSOTORDRO.Item("ORDR_COMMENT") = "Qty Ordered"
            dst.Tables("SOTORDRO").Rows.Add(rowSOTORDRO)

            rowSOTORDRO = dst.Tables("SOTORDRO").NewRow
            rowSOTORDRO.Item("ORDR_REL_HOLD_CODES") = InvalidItem
            rowSOTORDRO.Item("ORDR_COMMENT") = "Invalid Item"
            dst.Tables("SOTORDRO").Rows.Add(rowSOTORDRO)

            rowSOTORDRO = dst.Tables("SOTORDRO").NewRow
            rowSOTORDRO.Item("ORDR_REL_HOLD_CODES") = InvalidUOM
            rowSOTORDRO.Item("ORDR_COMMENT") = "Invalid UOM"
            dst.Tables("SOTORDRO").Rows.Add(rowSOTORDRO)

            rowSOTORDRO = dst.Tables("SOTORDRO").NewRow
            rowSOTORDRO.Item("ORDR_REL_HOLD_CODES") = FrozenInactiveItem
            rowSOTORDRO.Item("ORDR_COMMENT") = "Frozen / Inactive Item"
            dst.Tables("SOTORDRO").Rows.Add(rowSOTORDRO)

            rowSOTORDRO = dst.Tables("SOTORDRO").NewRow
            rowSOTORDRO.Item("ORDR_REL_HOLD_CODES") = ItemAuthorizationError
            rowSOTORDRO.Item("ORDR_COMMENT") = "Item Authorization Error"
            dst.Tables("SOTORDRO").Rows.Add(rowSOTORDRO)

            rowSOTORDRO = dst.Tables("SOTORDRO").NewRow
            rowSOTORDRO.Item("ORDR_REL_HOLD_CODES") = RevenueItemNoPrice
            rowSOTORDRO.Item("ORDR_COMMENT") = "Revenue at 0 cost"
            dst.Tables("SOTORDRO").Rows.Add(rowSOTORDRO)

        End Sub

        Private Function FormatZipCode(ByVal zipCode As String) As String

            Dim strVal As String = zipCode

            If strVal.Length = 9 AndAlso strVal.EndsWith("0000") Then
                strVal = strVal.Substring(0, 5)
            End If

            If strVal.Length >= 5 And strVal.Length <> 9 Then
                strVal = strVal.Substring(0, 5)
            End If

            If strVal.Length = 9 Then
                strVal = strVal.Substring(0, 5) & "-" & strVal.Substring(5)
            End If

            strVal = strVal.Replace("'", "")

            Return strVal

        End Function

        Private Function FormatTelePhone(ByVal Telephone As String) As String

            Dim strval As String = String.Empty

            For Each Chr As Char In Telephone
                If Char.IsDigit(Chr) Then
                    strval &= Chr
                End If
            Next

            Return strval
        End Function
#End Region

#Region "Log Procedures"

        Private Function OpenLogFile() As Boolean

            Try

                Dim svcConfig As New ServiceConfig

                filefolder = svcConfig.FileFolder

                If Not My.Computer.FileSystem.DirectoryExists(filefolder) Then
                    My.Computer.FileSystem.CreateDirectory(filefolder)
                End If

                logFilename = Format(Now, "yyyyMMdd") & ".log"
                If logStreamWriter IsNot Nothing Then
                    logStreamWriter.Close()
                    logStreamWriter.Dispose()
                End If

                Dim logdirectory As String = filefolder
                If Not logdirectory.EndsWith("\") Then logdirectory &= "\"
                logdirectory &= "Logs\"

                logStreamWriter = New System.IO.StreamWriter(logdirectory & logFilename, True)

                If testMode Then
                    RecordLogEntry(Environment.NewLine)
                    RecordLogEntry(Environment.NewLine)
                    RecordLogEntry("Open Log File.")
                End If

                Return True

            Catch ex As Exception
                Return False
            End Try
        End Function

        Private Sub RecordLogEntry(ByVal message As String)

            If message = Environment.NewLine Then
                logStreamWriter.WriteLine(message)
            Else
                logStreamWriter.WriteLine(DateTime.Now & ": " & message)
            End If
        End Sub

        Public Sub CloseLog()
            If logStreamWriter IsNot Nothing Then
                logStreamWriter.Close()
                logStreamWriter.Dispose()
                logStreamWriter = Nothing
            End If
        End Sub

        Private Sub emailErrors(ByRef ORDR_SOURCE As String, ByVal numErrors As Int16)
            emailErrors(ORDR_SOURCE, numErrors, String.Empty, String.Empty)
        End Sub

        Private Sub emailErrors(ByRef ORDR_SOURCE As String, ByVal numErrors As Int16, ByVal note As String)
            emailErrors(ORDR_SOURCE, numErrors, note, String.Empty)
        End Sub

        Private Sub emailErrors(ByRef ORDR_SOURCE As String, ByVal numErrors As Int16, ByVal note As String, ByVal NoteOverride As String)
            Try
                Dim clsASTNOTE1 As TAC.ASCNOTE1

                If NoteOverride.Length > 0 Then
                    clsASTNOTE1 = New TAC.ASCNOTE1(NoteOverride, dst, "")
                Else
                    clsASTNOTE1 = New TAC.ASCNOTE1("IMPERROR_" & ORDR_SOURCE, dst, "")
                End If

                If note.Length = 0 Then
                    note = "There were " & numErrors & " import error(s) on " & DateTime.Now.ToLongDateString & " " & DateTime.Now.ToLongTimeString
                End If

                clsASTNOTE1.Note = note
                clsASTNOTE1.CreateComponents()
                clsASTNOTE1.EmailDocument()
            Catch ex As Exception
                RecordLogEntry("Send Email: " & ex.Message)
            End Try
        End Sub

#End Region

    End Class

End Namespace


