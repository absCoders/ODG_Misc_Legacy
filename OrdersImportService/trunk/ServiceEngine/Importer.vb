Imports ServiceEngine.Extensions

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

        Private Const testMode As Boolean = False
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

        'Detail Errors
        Private Const QtyOrdered = "Q"
        Private Const InvalidItem = "I"
        Private Const InvalidUOM = "U"
        Private Const FrozenInactiveItem = "F"
        Private Const ItemAuthorizationError = "0"
        Private Const RevenueItemNoPrice = "N"

        Private timerPeriod As Integer = 10

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
                importInProcess = True

                If Not OpenLogFile() Then
                    importInProcess = False
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
                            ProcessSalesOrders()
                        End If
                    End If
                End If

                If testMode Then RecordLogEntry("Exit MainProcess.")
                RecordLogEntry("Closing Log file.")
                CloseLog()

            Catch ex As Exception
                RecordLogEntry("MainProcess: " & ex.Message)
            Finally
                importInProcess = False
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

                If ABSolution.ASCMAIN1.DBS_PASSWORD = "" OrElse ABSolution.ASCMAIN1.DBS_PASSWORD = "" OrElse ABSolution.ASCMAIN1.DBS_SERVER = "" Then
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

                LogIntoDatabase = True

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

        Private Sub ProcessSalesOrders()

            Try

                If testMode Then RecordLogEntry("Enter ProcessSalesOrders.")

                Dim ORDR_SOURCE As String = String.Empty

                ABSolution.ASCMAIN1.SESSION_NO = ABSolution.ASCMAIN1.Next_Control_No("ASTLOGS1.SESSION_NO", 1)

                If ABSolution.ASCMAIN1.ActiveForm Is Nothing Then
                    ABSolution.ASCMAIN1.ActiveForm = New ABSolution.ASFBASE1
                End If
                ABSolution.ASCMAIN1.ActiveForm.SELECTION_NO = "1"

                For Each rowXMTXREF1 As DataRow In ABSolution.ASCDATA1.GetDataTable("SELECT DISTINCT ORDR_SOURCE FROM XMTXREF1 WHERE ORDR_SOURCE IS NOT NULL").Rows

                    ORDR_SOURCE = rowXMTXREF1.Item("ORDR_SOURCE") & String.Empty

                    If Not ABSolution.ASCMAIN1.Logical_Lock("IMPSVC01", ORDR_SOURCE, False, False, True, 1) Then
                        RecordLogEntry("Order Import Type: " & ORDR_SOURCE & " locked by previous instance.")
                        Continue For
                    End If

                    If Not ClearDataSetTables(True) Then
                        Continue For
                    End If

                    ImportErrorNotification = New Hashtable

                    Select Case ORDR_SOURCE
                        Case "X" ' Web Service
                            ProcessWebServiceSalesOrders(ORDR_SOURCE)
                    End Select

                    ABSolution.ASCMAIN1.MultiTask_Release(, , 1)

                    If ImportErrorNotification.Keys.Count > 0 Then
                        For Each Item As DictionaryEntry In ImportErrorNotification
                            emailErrors(Item.Key, Item.Value)
                        Next
                    End If
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

            Try
                If testMode Then RecordLogEntry("Enter ProcessWebServiceSalesOrders.")

                baseClass.clsASCBASE1.Fill_Records("XSTORDR1", String.Empty, True, "SELECT * FROM XSTORDR1 WHERE NVL(PROCESS_IND, '0') = '0'")
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
                        rowSOTORDRX.Item("ITEM_BASE_CURVE") = rowXSTORDR2.Item("ITEM_BASE_CURVE") & String.Empty
                        rowSOTORDRX.Item("ITEM_SPHERE_POWER") = rowXSTORDR2.Item("ITEM_SPHERE_POWER") & String.Empty
                        rowSOTORDRX.Item("ITEM_CYLINDER") = rowXSTORDR2.Item("ITEM_CYLINDER") & String.Empty
                        rowSOTORDRX.Item("ITEM_AXIS") = rowXSTORDR2.Item("ITEM_AXIS") & String.Empty
                        rowSOTORDRX.Item("ITEM_DIAMETER") = rowXSTORDR2.Item("ITEM_DIAMETER") & String.Empty
                        rowSOTORDRX.Item("ITEM_ADD_POWER") = rowXSTORDR2.Item("ITEM_ADD_POWER") & String.Empty
                        rowSOTORDRX.Item("ITEM_COLOR") = rowXSTORDR2.Item("ITEM_COLOR") & String.Empty
                        rowSOTORDRX.Item("ITEM_MULTIFOCAL") = rowXSTORDR2.Item("ITEM_MULTIFOCAL") & String.Empty
                        rowSOTORDRX.Item("ITEM_NOTE") = rowXSTORDR2.Item("ITEM_NOTE") & String.Empty
                        rowSOTORDRX.Item("NEAR_DISTANCE") = String.Empty

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
                        rowSOTORDRX.Item("PAYMENT_METHOD") = rowXSTORDR1.Item("PAYMENT_METHOD") & String.Empty

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

                        rowSOTORDRX.Item("PRESCRIBING_DOCTOR") = rowXSTORDR1.Item("PRESCRIBING_DOCTOR") & String.Empty
                        rowSOTORDRX.Item("TAX_SHIPPING") = rowXSTORDR1.Item("TAX_SHIPPING") & String.Empty
                        rowSOTORDRX.Item("PATIENT_STAX_RATE") = rowXSTORDR1.Item("PATIENT_STAX_RATE") & String.Empty
                        rowSOTORDRX.Item("OFFICE_WEBSITE") = rowXSTORDR1.Item("OFFICE_WEBSITE") & String.Empty

                        dst.Tables("SOTORDRX").Rows.Add(rowSOTORDRX)
                    Next

                    ORDR_NO = String.Empty
                    If CreateSalesOrder(ORDR_NO, CREATE_SHIP_TO, SELECT_SHIP_TO_BY_TELE, ORDR_LINE_SOURCE, ORDR_SOURCE, CALLER_NAME) Then
                        ' All orders are ship complete
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

        Private Function CreateOrderShipTo(ByVal ORDR_NO As String, ByRef rowSOTORDRX As DataRow, ByVal shipToPatient As Boolean) As Boolean

            Dim rowSOTORDR5 As DataRow = Nothing

            Try

                If testMode Then RecordLogEntry("Enter CreateOrderShipTo: " & ORDR_NO)

                rowSOTORDR5 = dst.Tables("SOTORDR5").NewRow
                rowSOTORDR5.Item("ORDR_NO") = ORDR_NO
                rowSOTORDR5.Item("CUST_ADDR_TYPE") = "ST"
                dst.Tables("SOTORDR5").Rows.Add(rowSOTORDR5)

                If shipToPatient Then
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
                    rowSOTORDR5.Item("CUST_FAX") = String.Empty
                    rowSOTORDR5.Item("CUST_EMAIL") = String.Empty

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
                                          ByVal ORDR_SOURCE As String, ByVal CALLER_NAME As String) As Boolean

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

                rowSOTSVIA1 = baseClass.LookUp("SOTSVIA1", (rowSOTORDRX.Item("SHIP_VIA_CODE") & String.Empty).ToString.Trim)

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
                If Not Me.SetBillToAttributes(CUST_CODE, CUST_SHIP_TO_NO, rowSOTORDR1, errorCodes) Then
                    errorCodes &= "K"
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

                rowSOTORDR1.Item("ORDR_CUST_PO") = rowSOTORDRX.Item("PO_NO") & String.Empty
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

                rowSOTORDR1.Item("FRT_CONT_NO") = 0
                rowSOTORDR1.Item("PATIENT_NO") = String.Empty
                rowSOTORDR1.Item("ORDR_COMMENT") = String.Empty
                rowSOTORDR1.Item("EDI_CUST_REF_NO") = rowSOTORDRX.Item("EDI_CUST_REF_NO") & String.Empty
                rowSOTORDR1.Item("ORDR_CALLER_NAME") = CALLER_NAME

                Dim ORDR_COMMENT As String = (rowSOTORDRX.Item("ORDR_COMMENT") & String.Empty).ToString.Trim
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
                    rowSOTORDR2.Item("ORDR_LR") = rowImportDetails.Item("ORDR_LR")

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

                If Not CreateOrderShipTo(ORDR_NO, rowSOTORDRX, shipToPatient) Then
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
                    rowSOTORDR1.Item("ORDR_STATUS_WEB") = rowSOTORDR1.Item("ORDR_SOURCE") & String.Empty
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
                If (rowSOTORDR1.Item("ORDR_LOCK_SHIP_VIA") & String.Empty) = "1" Then Return True

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

                rowSOTORDR1.Item("CUST_NAME") = String.Empty
                rowSOTORDR1.Item("CUST_BILL_TO_CUST") = String.Empty

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
                    rowSOTORDR1.Item("CUST_NAME") = "Unknown Customer"
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
        ''' Sets Order Detail Information based on the Item's Attributes
        ''' </summary>
        ''' <param name="rowSOTORDR2"></param>
        ''' <param name="errorCodes"></param>
        ''' <remarks></remarks>
        Private Sub SetItemInfo(ByRef rowSOTORDR2 As DataRow, ByRef errorCodes As String)
            Try

                If testMode Then RecordLogEntry("Enter SetItemInfo.")

                Dim ITEM_CODE As String = rowSOTORDR2.Item("ITEM_CODE") & String.Empty
                Dim ITEM_DESC2 As String = rowSOTORDR2.Item("ITEM_DESC2") & String.Empty
                Dim ITEM_DESC2_X() As String = ITEM_DESC2.Split("/")

                rowICTITEM1 = Nothing

                ' See if the item exists in the item master
                If ITEM_CODE.Length > 0 Then
                    rowICTITEM1 = baseClass.LookUp("ICTITEM1", ITEM_CODE)
                End If

                ' If the item is not in the Item master then look at the catalogue
                If rowICTITEM1 Is Nothing Then
                    rowICTITEM1 = baseClass.LookUp("ICTCATL1", ITEM_CODE)
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

                    If rowARTCUST2.Item("CUST_SHIP_TO_STATUS") & String.Empty = "C" Or _
                        rowARTCUST2.Item("CUST_SHIP_TO_ORDER_BLOCK") & String.Empty = "1" Then
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

                    If ClearXMTtables Then
                        .Tables("XSTORDR1").Clear()
                        .Tables("XSTORDR2").Clear()
                    End If
                End With

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

                    With .Tables("SOTORDR2")
                        .Columns.Add("ORDR_LNO_EXT", GetType(System.Double), "ISNULL(ORDR_QTY, 0) * ISNULL(ORDR_UNIT_PRICE, 0)")
                    End With

                    baseClass.Create_TDA(.Tables.Add, "SOTORDR3", "*", 1)
                    baseClass.Create_TDA(.Tables.Add, "SOTORDR5", "*", 1)

                    baseClass.Create_TDA(.Tables.Add, "SOTORDRW", "*", 1)
                    baseClass.Create_TDA(.Tables.Add, "SOTORDRE", "*", 0)

                    baseClass.Create_TDA(.Tables.Add, "XSTORDR1", "*", 1)
                    baseClass.Create_TDA(.Tables.Add, "XSTORDR2", "*", 1)

                    baseClass.Create_TDA(.Tables.Add, "XMTXREF1", "*")
                    baseClass.Fill_Records("XMTXREF1", String.Empty, True, "Select * From XMTXREF1")

                    baseClass.Create_TDA(dst.Tables.Add, "SOTORDRO", "Select LPAD( ' ', 15) ORDR_REL_HOLD_CODES, ORDR_COMMENT From SOTORDR1", , False)

                    LoadTablesForPricing()

                    baseClass.Create_TDA(.Tables.Add, "PPTPARM1", "*")
                    baseClass.Fill_Records("PPTPARM1", String.Empty, True, "Select * From PPTPARM1")

                    Dim rowSOTPARM1 As DataRow = ABSolution.ASCDATA1.GetDataRow("SELECT * FROM SOTPARM1 WHERE SO_PARM_KEY = 'Z'")

                    If rowSOTPARM1 IsNot Nothing Then
                        DpdDefaultShipViaCode = (rowSOTPARM1.Item("SO_PARM_SHIP_VIA_CODE_DPD") & String.Empty).ToString.Trim
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

                    STAX_CODE_states = New List(Of String)
                    sql = "Select Distinct STATE from ARTSTAX2"
                    For Each row As DataRow In ABSolution.ASCDATA1.GetDataTable(sql).Rows
                        STAX_CODE_states.Add(row.Item(0))
                    Next

                    ' Create a work table to stuff the data into, this way there is only one
                    ' common routine to create the sales order
                    .Tables.Add("SOTORDRX")
                    .Tables("SOTORDRX").Columns.Add("CUST_CODE", GetType(System.String))
                    .Tables("SOTORDRX").Columns.Add("CUST_SHIP_TO_NO", GetType(System.String))
                    .Tables("SOTORDRX").Columns.Add("ORDR_DATE", GetType(System.String))
                    .Tables("SOTORDRX").Columns.Add("ORDR_DPD", GetType(System.String))
                    .Tables("SOTORDRX").Columns.Add("SHIP_VIA_CODE", GetType(System.String))
                    .Tables("SOTORDRX").Columns.Add("PATIENT_NAME", GetType(System.String))
                    .Tables("SOTORDRX").Columns.Add("EDI_CUST_REF_NO", GetType(System.String))
                    .Tables("SOTORDRX").Columns.Add("ORDR_SOURCE", GetType(System.String))
                    .Tables("SOTORDRX").Columns.Add("ORDR_TYPE_CODE", GetType(System.String))

                    .Tables("SOTORDRX").Columns.Add("ORDR_NO", GetType(System.String))
                    .Tables("SOTORDRX").Columns.Add("ORDR_LNO", GetType(System.String))
                    .Tables("SOTORDRX").Columns.Add("CUST_LINE_REF", GetType(System.String))
                    .Tables("SOTORDRX").Columns.Add("ORDR_QTY", GetType(System.String))
                    .Tables("SOTORDRX").Columns.Add("ORDR_UNIT_PRICE_PATIENT", GetType(System.String))
                    .Tables("SOTORDRX").Columns.Add("ORDR_LR", GetType(System.String))
                    .Tables("SOTORDRX").Columns.Add("ORDR_LINE_SOURCE", GetType(System.String))
                    .Tables("SOTORDRX").Columns.Add("ITEM_CODE", GetType(System.String))

                    .Tables("SOTORDRX").Columns.Add("PRICE_CATGY_CODE", GetType(System.String))
                    .Tables("SOTORDRX").Columns.Add("ITEM_DESC", GetType(System.String))
                    .Tables("SOTORDRX").Columns.Add("ITEM_PROD_ID", GetType(System.String))
                    .Tables("SOTORDRX").Columns.Add("ITEM_BASE_CURVE", GetType(System.String))
                    .Tables("SOTORDRX").Columns.Add("ITEM_SPHERE_POWER", GetType(System.String))
                    .Tables("SOTORDRX").Columns.Add("ITEM_CYLINDER", GetType(System.String))
                    .Tables("SOTORDRX").Columns.Add("ITEM_AXIS", GetType(System.String))
                    .Tables("SOTORDRX").Columns.Add("ITEM_DIAMETER", GetType(System.String))
                    .Tables("SOTORDRX").Columns.Add("ITEM_ADD_POWER", GetType(System.String))
                    .Tables("SOTORDRX").Columns.Add("ITEM_COLOR", GetType(System.String))
                    .Tables("SOTORDRX").Columns.Add("NEAR_DISTANCE", GetType(System.String))
                    .Tables("SOTORDRX").Columns.Add("ITEM_MULTIFOCAL", GetType(System.String))
                    .Tables("SOTORDRX").Columns.Add("ITEM_NOTE", GetType(System.String))

                    .Tables("SOTORDRX").Columns.Add("CUST_SHIP_TO_NAME", GetType(System.String))
                    .Tables("SOTORDRX").Columns.Add("CUST_SHIP_TO_ADDR1", GetType(System.String))
                    .Tables("SOTORDRX").Columns.Add("CUST_SHIP_TO_ADDR2", GetType(System.String))
                    .Tables("SOTORDRX").Columns.Add("CUST_SHIP_TO_ADDR3", GetType(System.String))
                    .Tables("SOTORDRX").Columns.Add("CUST_SHIP_TO_CITY", GetType(System.String))
                    .Tables("SOTORDRX").Columns.Add("CUST_SHIP_TO_STATE", GetType(System.String))
                    .Tables("SOTORDRX").Columns.Add("CUST_SHIP_TO_ZIP_CODE", GetType(System.String))
                    .Tables("SOTORDRX").Columns.Add("CUST_SHIP_TO_PHONE", GetType(System.String))
                    .Tables("SOTORDRX").Columns.Add("CUST_SHIP_TO_COUNTRY", GetType(System.String))

                    .Tables("SOTORDRX").Columns.Add("BILLING_NAME", GetType(System.String))
                    .Tables("SOTORDRX").Columns.Add("BILLING_ADDRESS1", GetType(System.String))
                    .Tables("SOTORDRX").Columns.Add("BILLING_ADDRESS2", GetType(System.String))
                    .Tables("SOTORDRX").Columns.Add("BILLING_CITY", GetType(System.String))
                    .Tables("SOTORDRX").Columns.Add("BILLING_STATE", GetType(System.String))
                    .Tables("SOTORDRX").Columns.Add("BILLING_ZIP", GetType(System.String))
                    .Tables("SOTORDRX").Columns.Add("PAYMENT_METHOD", GetType(System.String))

                    .Tables("SOTORDRX").Columns.Add("CUST_NAME", GetType(System.String))
                    .Tables("SOTORDRX").Columns.Add("CUST_ADDR1", GetType(System.String))
                    .Tables("SOTORDRX").Columns.Add("CUST_ADDR2", GetType(System.String))
                    .Tables("SOTORDRX").Columns.Add("CUST_CITY", GetType(System.String))
                    .Tables("SOTORDRX").Columns.Add("CUST_STATE", GetType(System.String))
                    .Tables("SOTORDRX").Columns.Add("CUST_ZIP_CODE", GetType(System.String))
                    .Tables("SOTORDRX").Columns.Add("CUST_PHONE", GetType(System.String))

                    .Tables("SOTORDRX").Columns.Add("ITEM_DESC2", GetType(System.String))
                    .Tables("SOTORDRX").Columns.Add("ORDR_LOCK_SHIP_VIA", GetType(System.String))
                    .Tables("SOTORDRX").Columns.Add("ORDR_COMMENT", GetType(System.String))

                    .Tables("SOTORDRX").Columns.Add("PRESCRIBING_DOCTOR", GetType(System.String))
                    .Tables("SOTORDRX").Columns.Add("TAX_SHIPPING", GetType(System.String))
                    .Tables("SOTORDRX").Columns.Add("PATIENT_STAX_RATE", GetType(System.String))
                    .Tables("SOTORDRX").Columns.Add("OFFICE_WEBSITE", GetType(System.String))

                    .Tables("SOTORDRX").Columns.Add("INIT_DATE", GetType(System.DateTime))
                    .Tables("SOTORDRX").Columns.Add("PO_NO", GetType(System.String))
                    .Tables("SOTORDRX").Columns.Add("OFFICE_MANAGER", GetType(System.String))
                    .Tables("SOTORDRX").Columns.Add("PROMO_CODE", GetType(System.String))
                    .Tables("SOTORDRX").Columns.Add("PATIENT_DISCOUNT_AMOUNT", GetType(System.String))
                    .Tables("SOTORDRX").Columns.Add("PATIENT_ID", GetType(System.String))

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

                If testMode Then RecordLogEntry(Environment.NewLine & Environment.NewLine & "Open Log File.")

                Return True

            Catch ex As Exception
                Return False
            End Try
        End Function

        Private Sub RecordLogEntry(ByVal message As String)
            logStreamWriter.WriteLine(DateTime.Now & ": " & message)
        End Sub

        Public Sub CloseLog()
            If logStreamWriter IsNot Nothing Then
                logStreamWriter.Close()
                logStreamWriter.Dispose()
                logStreamWriter = Nothing
            End If
        End Sub

        Private Sub emailErrors(ByRef ORDR_SOURCE As String, ByVal numErrors As Int16)
            Try

                Dim clsASTNOTE1 As New TAC.ASCNOTE1("IMPERROR_" & ORDR_SOURCE, dst, "")
                Dim note As String = "There were " & numErrors & " import error(s) on " & DateTime.Now.ToLongDateString & " " & DateTime.Now.ToLongTimeString
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


