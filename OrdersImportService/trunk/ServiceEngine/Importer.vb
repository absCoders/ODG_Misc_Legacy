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

        Private rowARTCUST1 As DataRow = Nothing
        Private rowARTCUST2 As DataRow = Nothing
        Private rowARTCUST3 As DataRow = Nothing

        Private rowSOTSVIA1 As DataRow = Nothing
        Private rowTATTERM1 As DataRow = Nothing
        Private rowICTITEM1 As DataRow = Nothing

        Private dst As DataSet

        Private Const testMode As Boolean = True
        Private ImportErrorNotification As Hashtable

        ' Header Errors
        Private Const InvalidSoldTo = "B"
        Private Const InvalidShipTo = "S"
        Private Const InvalidShipVia = "V"
        Private Const ReviewSpecInstr = "R"
        Private Const InvalidTaxCode = "X"
        Private Const InvalidTermsCode = "T"
        Private Const InvalidSalesTax = "G"
        Private Const InvalidDPD = "D"
        Private Const InvalidPricing = "H"
        Private Const InvalidSalesOrderTotal = "J"
        Private Const ShipToClosed = "C"
        Private Const ShipToOrderBlocked = "A"

        'Detail Errors
        Private Const QtyOrdered = "Q"
        Private Const InvalidItem = "I"
        Private Const InvalidUOM = "U"
        Private Const FrozenInactiveItem = "F"
        Private Const ItemAuthorizationError = "0"
        Private Const RevenueItemNoPrice = "N"

        Private timerPeriod As Integer = 10

        Private ImportedFiles As List(Of String) = New List(Of String)
        Private ftpFileList As List(Of String) = New List(Of String)

        Private WithEvents Ftp1 As New nsoftware.IPWorks.Ftp

        Private SO_PARM_SHIP_ND As String = String.Empty
        Private SO_PARM_SHIP_ND_DPD As String = String.Empty
        Private SO_PARM_SHIP_ND_COD As String = String.Empty

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
                Dim DriveLetter As String = svcConfig.DriveLetter.ToString.ToUpper
                Dim DriveLetterIP As String = svcConfig.DriveLetterIP.ToString.ToUpper
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
            importTimer = New System.Threading.Timer _
                (New System.Threading.TimerCallback(AddressOf MainProcess), Nothing, 3000, timerPeriod * 1000 * 60) ' every period Minutes
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

                Dim ORDR_SOURCE As String = String.Empty

                ABSolution.ASCMAIN1.SESSION_NO = ABSolution.ASCMAIN1.Next_Control_No("ASTLOGS1.SESSION_NO", 1)

                If ABSolution.ASCMAIN1.ActiveForm Is Nothing Then
                    ABSolution.ASCMAIN1.ActiveForm = New ABSolution.ASFBASE1
                End If
                ABSolution.ASCMAIN1.ActiveForm.SELECTION_NO = "1"

                For Each rowSOTPARMP As DataRow In ABSolution.ASCDATA1.GetDataTable("SELECT DISTINCT SO_PARM_KEY ORDR_SOURCE FROM SOTPARMP WHERE NVL(SO_PARM_USE_SERVICE, '0') = '1' ORDER BY SO_PARM_KEY").Rows

                    ORDR_SOURCE = rowSOTPARMP.Item("ORDR_SOURCE") & String.Empty

                    ' House Keeping, common things to do
                    Ftp1 = New nsoftware.IPWorks.Ftp
                    System.Threading.Thread.Sleep(1000)
                    Ftp1.RuntimeLicense = nSoftwareftpkey

                    ftpFileList.Clear()
                    ImportedFiles.Clear()
                    baseClass.Fill_Records("SOTSVIAF", ORDR_SOURCE)

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
                        Case "X", "Y" ' Web Service - (X) 800 Anylens, (Y) Eyeconic
                            ProcessWebServiceSalesOrders(ORDR_SOURCE)
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
                    End Select

                    ABSolution.ASCMAIN1.MultiTask_Release(, , 1)
                    Ftp1.Dispose()

                    If ImportErrorNotification.Keys.Count > 0 Then
                        For Each Item As DictionaryEntry In ImportErrorNotification
                            emailErrors(Item.Key, Item.Value)
                        Next
                    End If

                    ImportErrorNotification.Clear()
                Next

                If testMode Then RecordLogEntry("Exit ProcessSalesOrders.")

            Catch ex As Exception
                RecordLogEntry("ProcessSalesOrders: " & ex.Message)
            End Try

        End Sub

        Private Sub ProcessShellSalesOrders(ByVal ORDER_SOURCE_CODE As String)

            Try
                If testMode Then RecordLogEntry("Enter ProcessShellSalesOrders.")

                ' Place loop here to process the sales orders

                If testMode Then RecordLogEntry("Exit ProcessShellSalesOrders.")

                RecordLogEntry("0 Sales Orders to rocess")

            Catch ex As Exception
                RecordLogEntry("ProcessShellSalesOrders: " & ex.Message)
            End Try
        End Sub

        Private Sub ProcessWebServiceSalesOrders(ByVal ORDR_SOURCE As String)

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

            Try
                If testMode Then RecordLogEntry("Enter ProcessWebServiceSalesOrders.")

                baseClass.clsASCBASE1.Fill_Records("SOTPARMC", ORDR_SOURCE)
                sql = String.Empty
                If dst.Tables("SOTPARMC").Rows.Count > 0 Then
                    sql = String.Empty
                    For Each rowSOTPARMC As DataRow In dst.Tables("SOTPARMC").Rows
                        sql &= ",'" & rowSOTPARMC.Item("CUST_CODE") & "'"
                    Next
                    If sql.Length > 0 Then
                        sql = sql.Substring(1)
                        sql = " And CUSTOMER_ID IN (" & sql & ")"
                    End If
                End If

                baseClass.clsASCBASE1.Fill_Records("XSTORDR1", String.Empty, True, "SELECT * FROM XSTORDR1 WHERE NVL(PROCESS_IND, '0') = '0'" & sql)
                If dst.Tables("XSTORDR1").Rows.Count = 0 Then
                    RecordLogEntry("0 Web Service Sales Orders to process.")
                    Exit Sub
                End If

                RecordLogEntry(dst.Tables("XSTORDR1").Rows.Count & " Web Service Sales Orders to process.")

                For Each rowXSTORDR1 As DataRow In dst.Tables("XSTORDR1").Select("", "XS_DOC_SEQ_NO")
                    ClearDataSetTables(False)
                    ORDR_LNO = 0
                    XML_ORDR_SOURCE = rowXSTORDR1.Item("ORDR_SOURCE") & String.Empty

                    ' rowXSTORDR1.Item("ORDR_SOURCE") examples: VSP, AnyLens
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
            Finally
                If salesOrdersProcessed > 0 Then
                    RecordLogEntry(salesOrdersProcessed & " Web Service Sales Orders imported")
                End If
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

            Dim CUST_CODE As String = String.Empty
            Dim CUST_SHIP_TO_NO As String = String.Empty
            Dim ORDR_LR As String = String.Empty
            Dim custShip As String = String.Empty
            Dim salesOrdersProcessed As Integer = 0
            Dim ITEM_DESC2 As String = String.Empty

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
            Finally
                Ftp1.Logoff()
                Ftp1.Dispose()
            End Try

            Try
                For Each orderFile As String In My.Computer.FileSystem.GetFiles(ftpConnection.LocalInDir, FileIO.SearchOption.SearchTopLevelOnly, "*" & ftpConnection.Filename & "*.csv")

                    orderFileName = My.Computer.FileSystem.GetName(orderFile)
                    RecordLogEntry("Importing file: " & orderFileName)

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

                            rowSOTORDRX.Item("ORDR_SOURCE") = ORDR_SOURCE
                            rowSOTORDRX.Item("ORDR_NO") = orderElements(0)
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
                                Dim rowARTCUST3 As DataRow = baseClass.LookUp("ARTCUST3", CUST_CODE)
                                If rowARTCUST3 IsNot Nothing Then
                                    If rowSOTORDRX.Item("ORDR_DPD") = "1" Then
                                        rowSOTORDRX.Item("SHIP_VIA_CODE") = rowARTCUST3.Item("SHIP_VIA_CODE_DPD")
                                    Else
                                        rowSOTORDRX.Item("SHIP_VIA_CODE") = rowARTCUST3.Item("SHIP_VIA_CODE")
                                    End If
                                End If
                            End If

                            If rowSOTORDRX.Item("SHIP_VIA_CODE") & String.Empty = String.Empty Then
                                rowSOTORDRX.Item("SHIP_VIA_CODE") = "SD"
                            End If

                            ' Second part of Else - As per Maria id DPD the ship complete
                            rowSOTORDRX.Item("ORDR_SHIP_COMPLETE") = IIf(orderElements(33) = "1", "1", "0") OrElse IIf(orderElements(4) = "1", "1", "0")

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
                Dim ORDR_NO As String = String.Empty
                Dim ORDR_CALLER_NAME As String = String.Empty
                Dim ORDR_SHIP_COMPLETE As String = String.Empty
                dst.Tables("SOTORDRX").Rows.Clear()

                ' Need to process each order individually for pricing reasons; therefore
                ' need to move the datat to a temp data table and process each order individually
                For Each headers As DataRow In ABSolution.ASCDATA1.GetDataTable("SELECT DISTINCT ORDR_NO FROM SOTORDRX WHERE PROCESS_IND IS NULL AND ORDR_SOURCE = :PARM1", String.Empty, "V", New Object() {ORDR_SOURCE}).Rows
                    ClearDataSetTables(True)
                    ORDR_NO = headers.Item("ORDR_NO") & String.Empty
                    baseClass.Fill_Records("SOTORDRX", New Object() {ORDR_SOURCE, ORDR_NO})

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
            Finally
                If salesOrdersProcessed > 0 Then
                    RecordLogEntry(salesOrdersProcessed & " " & ftpConnection.ConnectionDescription & " Orders imported.")
                End If
            End Try


        End Sub

        ''' <summary>
        ''' Replaces ABSolution SOTORDRE
        ''' </summary>
        ''' <param name="ORDR_SOURCE"></param>
        ''' <remarks></remarks>
        Private Sub ProcessEDISalesOrders(ByVal ORDR_SOURCE As String)

            Dim sql As String = String.Empty

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
            baseClass.Fill_Records("EDT850I1D", String.Empty, False, sql)

            Try
                For Each rowEDT850I1D As DataRow In dst.Tables("EDT850I1D").Select("", "EDI_ISA_NO")

                    EDI_ISA_NO = rowEDT850I1D.Item("EDI_ISA_NO") & String.Empty

                    sql = "Select * From EDT850I1 Where EDI_ISA_NO = '" & EDI_ISA_NO & "' AND EDI_BATCH_NO IS NULL"
                    Call baseClass.Fill_Records("EDT850I1", String.Empty, False, sql)

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

                        Call baseClass.Fill_Records("EDT850I2", EDI_DOC_SEQ_NO)
                        Call baseClass.Fill_Records("EDT850I5", EDI_DOC_SEQ_NO)

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
                    baseClass.Fill_Records("SOTORDRX", New Object() {ORDR_SOURCE, ORDR_NO})

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
            Finally
                If salesOrdersProcessed > 0 Then
                    RecordLogEntry(salesOrdersProcessed & " " & ORDR_SOURCE & " Orders imported.")
                End If
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
                    baseClass.Fill_Records("SOTORDRX", New Object() {ORDR_SOURCE, ORDR_NO})

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
            Finally
                If salesOrdersProcessed > 0 Then
                    RecordLogEntry(salesOrdersProcessed & " " & ftpConnection.ConnectionDescription & " Orders imported.")
                End If

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

            Try
                ImportedFiles.Clear()

                Dim xsdFile As String = vwConnection.LocalInDir & "ContactLensOrder.XSD"

                If Not My.Computer.FileSystem.FileExists(xsdFile) Then
                    RecordLogEntry("ProcessVisionWebSalesOrders: " & xsdFile & " could not be found.")
                    Exit Sub
                End If

                ' Create the DataSet to read the schema into.
                Dim vwXmlDataset As New DataSet
                'Create a FileStream object with the file path and name.
                Dim myFileStream As System.IO.FileStream = New System.IO.FileStream(xsdFile, System.IO.FileMode.Open)
                'Create a new XmlTextReader object with the FileStream.
                Dim myXmlTextReader As System.Xml.XmlTextReader = New System.Xml.XmlTextReader(myFileStream)
                'Read the schema into the DataSet and close the reader.
                vwXmlDataset.ReadXmlSchema(myXmlTextReader)
                myXmlTextReader.Close()

                For Each orderFile As String In My.Computer.FileSystem.GetFiles(vwConnection.LocalInDir, FileIO.SearchOption.SearchTopLevelOnly, "*.xml")

                    For Each tbl As DataTable In vwXmlDataset.Tables
                        tbl.Rows.Clear()
                        tbl.BeginLoadData()
                    Next

                    vwXmlDataset.ReadXml(orderFile)

                    For Each tbl As DataTable In vwXmlDataset.Tables
                        tbl.EndLoadData()
                    Next

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

                            creationDate = (rowSoftType.Item("Creation_date") & String.Empty).ToString.Replace("T", String.Empty)
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
                                rowSOTORDRX.Item("ORDR_CUST_PO") = TruncateField(rowSoftType.Item("Id") & IIf(rowHeader.Item("PurchaseOrderNumber") & String.Empty <> String.Empty, " : " & rowHeader.Item("PurchaseOrderNumber"), String.Empty), "SOTORDRX", "ORDR_CUST_PO")
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
                                                        rowSOTORDRX.Item("CUST_SHIP_TO_ADDR1") = TruncateField((rowData.Item("Street_Number") & String.Empty & " " & rowData.Item("Street_Name") & String.Empty).ToString.Trim, "SOTORDRX", "CUST_SHIP_TO_ADDR1")
                                                        rowSOTORDRX.Item("CUST_SHIP_TO_ADDR2") = TruncateField(rowData.Item("Suite") & String.Empty, "SOTORDRX", "CUST_SHIP_TO_ADDR2")
                                                        rowSOTORDRX.Item("CUST_SHIP_TO_ADDR3") = String.Empty
                                                        rowSOTORDRX.Item("CUST_SHIP_TO_CITY") = TruncateField(rowData.Item("City") & String.Empty, "SOTORDRX", "CUST_SHIP_TO_CITY")
                                                        rowSOTORDRX.Item("CUST_SHIP_TO_STATE") = TruncateField(rowData.Item("State") & String.Empty, "SOTORDRX", "CUST_SHIP_TO_STATE")
                                                        rowSOTORDRX.Item("CUST_SHIP_TO_ZIP_CODE") = TruncateField(rowData.Item("ZipCode") & String.Empty, "SOTORDRX", "CUST_SHIP_TO_ZIP_CODE")
                                                        Telephone = rowData.Item("TEL") & String.Empty
                                                        Telephone = FormatTelePhone(Telephone)
                                                        rowSOTORDRX.Item("CUST_SHIP_TO_PHONE") = TruncateField(Telephone, "SOTORDRX", "CUST_SHIP_TO_PHONE")
                                                        rowSOTORDRX.Item("CUST_SHIP_TO_COUNTRY") = TruncateField(rowData.Item("Country") & String.Empty, "SOTORDRX", "CUST_SHIP_TO_COUNTRY")
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
                                                        rowSOTORDRX.Item("CUST_ADDR1") = TruncateField((rowData.Item("Street_Number") & String.Empty & " " & rowData.Item("Street_Name") & String.Empty).ToString.Trim, "SOTORDRX", "CUST_ADDR1")
                                                        rowSOTORDRX.Item("CUST_ADDR2") = TruncateField(rowData.Item("Suite") & String.Empty, "SOTORDRX", "CUST_ADDR2")
                                                        rowSOTORDRX.Item("CUST_CITY") = TruncateField(rowData.Item("City") & String.Empty, "SOTORDRX", "CUST_CITY")
                                                        rowSOTORDRX.Item("CUST_STATE") = TruncateField(rowData.Item("State") & String.Empty, "SOTORDRX", "CUST_STATE")
                                                        rowSOTORDRX.Item("CUST_ZIP_CODE") = TruncateField(rowData.Item("Zipcode") & String.Empty, "SOTORDRX", "CUST_ZIP_CODE")
                                                        Telephone = rowData.Item("TEL") & String.Empty
                                                        Telephone = FormatTelePhone(Telephone)
                                                        rowSOTORDRX.Item("CUST_PHONE") = TruncateField(Telephone, "SOTORDRX", "CUST_PHONE")
                                                        rowSOTORDRX.Item("CUST_FAX") = String.Empty
                                                        rowSOTORDRX.Item("CUST_EMAIL") = String.Empty
                                                        rowSOTORDRX.Item("CUST_COUNTRY") = TruncateField(rowData.Item("Country") & String.Empty, "SOTORDRX", "CUST_COUNTRY")

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
                                    AttentionTo = rowData.Item("AttentionTo") & String.Empty

                                    If vwXmlDataset.Tables("DELIVERY_METHOD").Select("DELIVERY_id = " & DELIVERY_Id).Length > 0 Then
                                        rowData = vwXmlDataset.Tables("DELIVERY_METHOD").Select("DELIVERY_id = " & DELIVERY_Id)(0)
                                        rowSOTORDRX.Item("SHIP_VIA_CODE") = rowData.Item("Name") & String.Empty

                                        If (rowSOTORDRX.Item("SHIP_VIA_CODE") & String.Empty).ToString.ToUpper = "STANDARD CONTRACT" Then
                                            rowSOTORDRX.Item("SHIP_VIA_CODE") = "STANDARD"
                                        End If

                                        If rowSOTORDRX.Item("SHIP_VIA_CODE") & String.Empty = String.Empty Then
                                            rowSOTORDRX.Item("SHIP_VIA_CODE") = rowData.Item("Id") & String.Empty
                                        End If
                                    End If

                                End If

                                ' If not DPD and Not Standard delivery, then lock the ship via
                                If rowSOTORDRX.Item("ORDR_DPD") = "1" _
                                AndAlso (rowSOTORDRX.Item("SHIP_VIA_CODE") & String.Empty).ToString.Trim.Length > 0 _
                                AndAlso rowSOTORDRX.Item("SHIP_VIA_CODE") & String.Empty <> "STANDARD" Then
                                    rowSOTORDRX.Item("ORDR_LOCK_SHIP_VIA") = "1"
                                End If

                                ' Get DPD Address
                                If DELIVERY_Id.Length > 0 AndAlso rowSOTORDRX.Item("ORDR_DPD") = "1" Then
                                    If vwXmlDataset.Tables("ADDRESS").Select("DELIVERY_Id = " & DELIVERY_Id).Length > 0 Then
                                        rowData = vwXmlDataset.Tables("ADDRESS").Select("DELIVERY_Id = " & DELIVERY_Id)(0)
                                        rowSOTORDRX.Item("CUST_NAME") = TruncateField(AttentionTo, "SOTORDRX", "CUST_NAME")
                                        rowSOTORDRX.Item("CUST_ADDR1") = TruncateField((rowData.Item("Street_Number") & String.Empty & " " & rowData.Item("Street_Name") & String.Empty).ToString.Trim, "SOTORDRX", "CUST_ADDR1")
                                        rowSOTORDRX.Item("CUST_ADDR2") = TruncateField(rowData.Item("Suite") & String.Empty, "SOTORDRX", "CUST_ADDR2")
                                        rowSOTORDRX.Item("CUST_CITY") = TruncateField(rowData.Item("City") & String.Empty, "SOTORDRX", "CUST_CITY")
                                        rowSOTORDRX.Item("CUST_STATE") = TruncateField(rowData.Item("State") & String.Empty, "SOTORDRX", "CUST_STATE")
                                        rowSOTORDRX.Item("CUST_ZIP_CODE") = TruncateField(rowData.Item("Zipcode") & String.Empty, "SOTORDRX", "CUST_ZIP_CODE")
                                        Telephone = rowData.Item("TEL") & String.Empty
                                        Telephone = FormatTelePhone(Telephone)
                                        rowSOTORDRX.Item("CUST_PHONE") = TruncateField(Telephone, "SOTORDRX", "CUST_PHONE")
                                        rowSOTORDRX.Item("CUST_FAX") = String.Empty
                                        rowSOTORDRX.Item("CUST_EMAIL") = String.Empty
                                        rowSOTORDRX.Item("CUST_COUNTRY") = TruncateField(rowData.Item("Country") & String.Empty, "SOTORDRX", "CUST_COUNTRY")
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
                                            rowSOTORDRX.Item("PATIENT_NAME") = TruncateField((rowData.Item("FirstName") & " " & rowData.Item("LastName")).ToString.Trim, "SOTORDRX", "PATIENT_NAME")
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
                    baseClass.Fill_Records("SOTORDRX", New Object() {ORDR_SOURCE, ORDR_NO})

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
            Finally
                If salesordersprocessed > 0 Then
                    RecordLogEntry(salesordersprocessed & " " & vwConnection.ConnectionDescription & " Orders imported.")
                End If

                ' Move Xml files to the archive directory
                For Each orderFile As String In ImportedFiles
                    My.Computer.FileSystem.MoveFile(orderFile, vwConnection.LocalInDirArchive & My.Computer.FileSystem.GetName(orderFile), True)
                Next

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

                baseClass.Fill_Records("XSTORDRQ", String.Empty, True, "Select * From XSTORDRQ WHERE ORDR_SOURCE = 'V'")

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
                    If rowARTCUST1.Item("CUST_PHONE") & String.Empty = CUST_SHIP_TO_PHONE Then
                        CUST_SHIP_TO_NO = String.Empty
                    Else
                        sql = "SELECT * From ARTCUST2 WHERE CUST_CODE = :PARM1 AND CUST_SHIP_TO_PHONE = :PARM2"
                        rowARTCUST2 = ABSolution.ASCDATA1.GetDataRow(sql, "VV", New Object() {CUST_CODE, CUST_SHIP_TO_PHONE})
                        If rowARTCUST2 Is Nothing Then
                            CUST_SHIP_TO_NO = "xxxxxx"
                        End If
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
                If ORDR_COMMENT.Length > 0 Then
                    Me.AddCharNoDups(ReviewSpecInstr, SOTORDR1ErrorCodes)

                    rowSOTORDR1.Item("ORDR_COMMENT") = TruncateField(ORDR_COMMENT, "SOTORDR1", "ORDR_COMMENT")
                    rowSOTORDR1.Item("REVIEW_ORDR_TEXT") = "1"

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

                If Not CreateSalesOrderTax(ORDR_NO) Then
                    Me.AddCharNoDups(InvalidSalesTax, SOTORDR1ErrorCodes)
                End If

                SOTORDR1ErrorCodes &= String.Empty
                rowSOTORDR1.Item("ORDR_REL_HOLD_CODES") = SOTORDR1ErrorCodes.Trim

                If (rowSOTORDR1.Item("ORDR_REL_HOLD_CODES") & String.Empty).ToString.Trim.Length = 0 Then
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

                If (rowSOTORDR1.Item("ORDR_REL_HOLD_CODES") & String.Empty).ToString.Trim.Length = 0 Then
                    If Not Me.UpdateSalesOrderTotal(ORDR_NO) Then
                        rowSOTORDR1.Item("ORDR_REL_HOLD_CODES") = rowSOTORDR1.Item("ORDR_REL_HOLD_CODES") & InvalidSalesOrderTotal
                    End If
                End If

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
                CreateSalesOrder = True

            Catch ex As Exception
                CreateSalesOrder = False
                RecordLogEntry("CreateSalesOrder: " & ex.Message)
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
                clsSOCORDR1.Price_and_Qty()

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
                    ElseIf rowARTCUST3 IsNot Nothing Then
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

#End Region

#Region "DataSet Functions"

        Private Function ClearDataSetTables(ByVal ClearXMTtables As Boolean) As Boolean

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

                    If ClearXMTtables Then
                        .Tables("XSTORDR1").Clear()
                        .Tables("XSTORDR2").Clear()
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
                    baseClass.Fill_Records("XMTXREF1", String.Empty, True, "Select * From XMTXREF1")

                    baseClass.Create_TDA(.Tables.Add, "XSTORDRQ", "*", 2)

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
                    baseClass.Fill_Records("SOTPARM1", "Z")

                    baseClass.Create_TDA(.Tables.Add, "SOTPARMB", "*")
                    baseClass.Fill_Records("SOTPARMB", "Z")

                    SO_PARM_SHIP_ND = (dst.Tables("SOTPARMB").Rows(0).Item("SO_PARM_SHIP_ND") & String.Empty).ToString.Trim
                    SO_PARM_SHIP_ND_DPD = (dst.Tables("SOTPARMB").Rows(0).Item("SO_PARM_SHIP_ND_DPD") & String.Empty).ToString.Trim
                    SO_PARM_SHIP_ND_COD = (dst.Tables("SOTPARMB").Rows(0).Item("SO_PARM_SHIP_ND_COD") & String.Empty).ToString.Trim

                    baseClass.Create_TDA(.Tables.Add, "SOTPARMC", "*", "1", False)

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
                    End If

                    .clsASCBASE1.Update_Record_TDA("XSTORDR1")
                    .clsASCBASE1.Update_Record_TDA("XSTORDR2")

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

        Private Sub emailErrors(ByRef ORDR_SOURCE As String, ByVal numErrors As Int16, Optional ByVal note As String = "")
            Try

                Dim clsASTNOTE1 As New TAC.ASCNOTE1("IMPERROR_" & ORDR_SOURCE, dst, "")
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


