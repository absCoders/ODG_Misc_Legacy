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

#End Region

#Region "Instaniate Service"

        Public Sub New()

            Dim svcConfig As New ServiceConfig

            filefolder = svcConfig.FileFolder
            OpenLogFile()

            ' Log into Oracle
            If LogIntoDatabase() Then
                InitializeSettings()
                PrepareDatasetEntries()
                ProcessSalesOrders()
            End If

            CloseLog()
        End Sub

#End Region

#Region "Data Management"

        Public Sub LogIn()
            importTimer = New System.Threading.Timer _
                (New System.Threading.TimerCallback(AddressOf StartingProcess), Nothing, 60000, 600000) ' every 10 mins 
        End Sub

        Private Sub StartingProcess()
            ' Do nothing. just a way to start the service
        End Sub

        Private Function LogIntoDatabase() As Boolean
            LogIntoDatabase = False

            Try
                Dim svcConfig As New ServiceConfig
                ABSolution.ASCMAIN1.DBS_COMPANY = svcConfig.UID
                ABSolution.ASCMAIN1.DBS_PASSWORD = svcConfig.PWD
                ABSolution.ASCMAIN1.DBS_SERVER = svcConfig.TNS

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

            Catch ex As Exception
                LogIntoDatabase = False
                RecordLogEntry("Error logging into Oracle: " & ex.Message)
            End Try

        End Function

        Private Sub InitializeSettings()

            Dim INIT_DATE As Date = DateTime.Now + ABSolution.ASCMAIN1.NowTSD

            baseClass = New ABSolution.ASFBASE1
            pricingClass = New ABSolution.TACMAIN1

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

        End Sub

        Public Function GetSessionId() As Int32
            Dim _currentProcess As Process = Process.GetCurrentProcess()
            Dim _processID As Int32 = _currentProcess.Id
            Dim _sessionID As Int32
            Dim _result As Boolean = ProcessIdToSessionId(_processID, _sessionID)
            Return _sessionID
        End Function

        Private Sub ProcessSalesOrders()

            Dim orderSourceHash As New Hashtable
            orderSourceHash.Add("X", "XML")
            orderSourceHash.Add("O", "OptiPort")
            orderSourceHash.Add("V", "Vision Web")
            orderSourceHash.Add("F", "eyeFinity")
            orderSourceHash.Add("C", "Customer Excel")
            orderSourceHash.Add("Y", "Eyeconic")

            ABSolution.ASCMAIN1.SESSION_NO = "123456"

            If ABSolution.ASCMAIN1.ActiveForm Is Nothing Then
                ABSolution.ASCMAIN1.ActiveForm = New ABSolution.ASFBASE1
            End If
            ABSolution.ASCMAIN1.ActiveForm.SELECTION_NO = "1"

            For Each orderSourceCode As String In orderSourceHash.Keys

                If Not ABSolution.ASCMAIN1.Logical_Lock("IMPSVC01", orderSourceCode, False, False, True, 1) Then
                    RecordLogEntry("Order Import Type: " & orderSourceHash(orderSourceCode) & " locked by previous instance.")
                    Continue For
                End If

                ClearDataSetTables(True)

                Select Case orderSourceCode
                    Case "X" ' XML
                        ProcessXmlSalesOrders(orderSourceCode)
                    Case "O" ' OptiPort
                        ProcessOptiPortSalesOrders(orderSourceCode)
                    Case "V" ' Vision Web
                        ProcessVisionWebSalesOrders(orderSourceCode)
                    Case "F" ' eyeFinity
                        ProcessEyefinitySalesOrders(orderSourceCode)
                    Case "C" ' Customer Excel
                        ProcessExcelFormatSalesOrders(orderSourceCode)
                    Case "Y" ' Eyeconic
                        ProcessEyeconicSalesOrders(orderSourceCode)
                End Select

                ABSolution.ASCMAIN1.MultiTask_Release(, , 1)
            Next
        End Sub

        ''' <summary>
        ''' Process Optiport sales orders
        ''' </summary>
        ''' <remarks></remarks>
        Private Sub ProcessOptiPortSalesOrders(ByVal ORDER_SOURCE_CODE As String)
        End Sub

        Private Sub ProcessXmlSalesOrders(ByVal ORDER_SOURCE_CODE As String)
        End Sub

        Private Sub ProcessVisionWebSalesOrders(ByVal ORDER_SOURCE_CODE As String)
        End Sub

        Private Sub ProcessEyefinitySalesOrders(ByVal ORDER_SOURCE_CODE As String)
        End Sub

        Private Sub ProcessExcelFormatSalesOrders(ByVal ORDER_SOURCE_CODE As String)
        End Sub

        Private Sub ProcessEyeconicSalesOrders(ByVal ORDER_SOURCE As String)

            Dim rowSOTORDRX As DataRow = Nothing
            Dim salesOrdersProcessed As Int16 = 0
            Dim ORDR_LINE_SOURCE As String = String.Empty
            Dim ORDR_NO As String = String.Empty
            Dim ORDR_LNO As Int16 = 0

            Try
                baseClass.clsASCBASE1.Fill_Records("XSTORDR1", String.Empty, True, "SELECT * FROM XSTORDR1 WHERE NVL(PROCESS_IND, '0') = '0' AND ORDR_SOURCE = 'VSP'")
                If dst.Tables("XSTORDR1").Rows.Count = 0 Then
                    RecordLogEntry("No Eyeconic Sales Orders to process.")
                    Exit Sub
                End If

                RecordLogEntry(dst.Tables("XSTORDR1").Rows.Count & " Eyeconic Sales Orders to process.")

                For Each rowXSTORDR1 As DataRow In dst.Tables("XSTORDR1").Select("", "XS_DOC_SEQ_NO")
                    ClearDataSetTables(False)
                    ORDR_LNO = 0

                    ' Flag entry as getting processed
                    rowXSTORDR1.Item("PROCESS_IND") = "1"

                    baseClass.clsASCBASE1.Fill_Records("XSTORDR2", (rowXSTORDR1.Item("XS_DOC_SEQ_NO") & String.Empty).ToString.Trim)

                    ' No details so get the hell out of here. We should charge $50.00 for processing and handling
                    If dst.Tables("XSTORDR2").Rows.Count = 0 Then
                        RecordLogEntry("No Eyeconic Sales Orders Details for XS Doc Seq No: " & (rowXSTORDR1.Item("XS_DOC_SEQ_NO") & String.Empty).ToString.Trim)
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
                        rowSOTORDRX.Item("ORDR_SOURCE") = ORDER_SOURCE
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

                        ' Convert Ordr_Line_Source to a code
                        ORDR_LINE_SOURCE = rowXSTORDR2.Item("ORDR_SOURCE") & String.Empty
                        If dst.Tables("XMTXREF1").Select("XML_ORDR_SOURCE = '" & ORDR_LINE_SOURCE & "'", "").Length > 0 Then
                            ORDR_LINE_SOURCE = dst.Tables("XMTXREF1").Select("XML_ORDR_SOURCE = '" & ORDR_LINE_SOURCE & "'", "")(0).Item("ORDR_LINE_SOURCE") & String.Empty
                        Else
                            ORDR_LINE_SOURCE = String.Empty
                        End If
                        rowSOTORDRX.Item("ORDR_LINE_SOURCE") = TruncateField(ORDR_LINE_SOURCE, "SOTORDR2", "ORDR_LINE_SOURCE")

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

                        rowSOTORDRX.Item("ORDR_LOCK_SHIP_VIA") = "1"
                        rowSOTORDRX.Item("ORDR_COMMENT") = String.Empty

                        rowSOTORDRX.Item("PRESCRIBING_DOCTOR") = rowXSTORDR1.Item("PRESCRIBING_DOCTOR") & String.Empty
                        rowSOTORDRX.Item("TAX_SHIPPING") = rowXSTORDR1.Item("TAX_SHIPPING") & String.Empty
                        rowSOTORDRX.Item("PATIENT_STAX_RATE") = rowXSTORDR1.Item("PATIENT_STAX_RATE") & String.Empty
                        rowSOTORDRX.Item("OFFICE_WEBSITE") = rowXSTORDR1.Item("OFFICE_WEBSITE") & String.Empty

                        dst.Tables("SOTORDRX").Rows.Add(rowSOTORDRX)
                    Next

                    ORDR_NO = String.Empty
                    If CreateSalesOrder(True, ORDR_NO) Then
                        ' All orders are ship complete
                        dst.Tables("SOTORDR1").Rows(0).Item("ORDR_SHIP_COMPLETE") = "1"
                        rowXSTORDR1.Item("ORDR_NO") = ORDR_NO
                        For Each rowXSTORDR2 As DataRow In dst.Tables("XSTORDR2").Rows
                            rowXSTORDR2.Item("ORDR_NO") = ORDR_NO
                        Next

                        UpdateDataSetTables()
                        salesOrdersProcessed += 1
                    Else
                        rowXSTORDR1.Item("PROCESS_IND") = "E"
                        RecordLogEntry("Eyeconic Doc Seq No: " & (rowXSTORDR1.Item("XS_DOC_SEQ_NO") & String.Empty).ToString & " not imported")

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

            Catch ex As Exception
                RecordLogEntry("ProcessEyeconicSalesOrders: " & ex.Message)
            Finally
                If salesOrdersProcessed > 0 Then
                    RecordLogEntry(salesOrdersProcessed & " Eyeconic Sales Orders imported")
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

        End Sub

        Private Sub CreateCustomerShipTo(ByVal CUST_CODE As String, ByVal CUST_SHIP_TO_NO As String, ByRef rowSOTORDRX As DataRow)

            Dim sql As String = String.Empty
            Dim CUST_SHIP_TO_SHIP_VIA_CODE As String = String.Empty
            Dim CUST_SHIP_TO_STATE As String = (rowSOTORDRX.Item("CUST_SHIP_TO_STATE") & String.Empty).ToString.Replace("'", "").Trim

            Dim rowSOTSVIAE As DataRow = ABSolution.ASCDATA1.GetDataRow("SELECT * FROM SOTSVIAE WHERE CUST_CODE = '" & CUST_CODE & "' AND STATE_CODE = '" & CUST_SHIP_TO_STATE & "'")
            If rowSOTSVIAE IsNot Nothing Then
                CUST_SHIP_TO_SHIP_VIA_CODE = (rowSOTSVIAE.Item("SHIP_VIA_CODE") & String.Empty).ToString.Trim
            End If

            sql = "INSERT INTO ARTCUST2 "
            sql &= " ("
            sql &= "CUST_CODE, CUST_SHIP_TO_NO, CUST_SHIP_TO_NAME,"
            sql &= " CUST_SHIP_TO_ADDR1, CUST_SHIP_TO_ADDR2, CUST_SHIP_TO_ADDR3,"
            sql &= " CUST_SHIP_TO_CITY, CUST_SHIP_TO_STATE, CUST_SHIP_TO_ZIP_CODE,"
            sql &= " CUST_SHIP_TO_COUNTRY, INIT_DATE, INIT_OPER,"
            sql &= " LAST_DATE, LAST_OPER, CUST_SHIP_TO_STATUS, CUST_SHIP_TO_SHIP_VIA_CODE"
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
            sql &= ")"

            ABSolution.ASCDATA1.ExecuteSQL(sql)
            System.Threading.Thread.Sleep(2000)
            rowARTCUST2 = baseClass.LookUp("ARTCUST2", New String() {CUST_CODE, CUST_SHIP_TO_NO})
        End Sub

        Private Sub CreateOrderBillTo(ByVal ORDR_NO As String)

            Dim rowSOTORDR5 As DataRow = dst.Tables("SOTORDR5").NewRow

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

        End Sub

        ''' <summary>
        ''' Creates and Entry in SOTORDRW for errors on the order
        ''' </summary>
        ''' <param name="ErrorCode"></param>
        ''' <param name="ORDR_LNO"></param>
        ''' <remarks></remarks>
        Private Sub CreateOrderErrorRecord(ByVal ORDR_NO As String, ByVal ORDR_LNO As Integer, ByVal ErrorCode As String, Optional ByVal ErrorMessage As String = "")
            Try
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

            Catch ex As Exception
                RecordLogEntry(ex.Message)
            End Try
        End Sub

        Private Sub CreateOrderPatientBillTo(ByVal ORDR_NO As String, ByRef rowSOTORDRX As DataRow)

            If (rowSOTORDRX.Item("BILLING_NAME") & String.Empty).ToString.Trim.Length = 0 _
                OrElse (rowSOTORDRX.Item("BILLING_ADDRESS1") & String.Empty).ToString.Trim.Length = 0 Then
                Exit Sub
            End If

            Dim rowSOTORDR5 As DataRow = dst.Tables("SOTORDR5").NewRow
            rowSOTORDR5.Item("ORDR_NO") = ORDR_NO
            rowSOTORDR5.Item("CUST_ADDR_TYPE") = "PB"
            dst.Tables("SOTORDR5").Rows.Add(rowSOTORDR5)

            rowSOTORDR5.Item("CUST_NAME") = rowSOTORDRX.Item("BILLING_NAME") & String.Empty
            rowSOTORDR5.Item("CUST_ADDR1") = rowSOTORDRX.Item("BILLING_ADDRESS1") & String.Empty
            rowSOTORDR5.Item("CUST_ADDR2") = rowSOTORDRX.Item("BILLING_ADDRESS2") & String.Empty
            rowSOTORDR5.Item("CUST_CITY") = rowSOTORDRX.Item("BILLING_CITY") & String.Empty
            rowSOTORDR5.Item("CUST_STATE") = rowSOTORDRX.Item("BILLING_STATE") & String.Empty
            rowSOTORDR5.Item("CUST_ZIP_CODE") = rowSOTORDRX.Item("BILLING_ZIP") & String.Empty
            rowSOTORDR5.Item("CUST_COUNTRY") = "US"
            rowSOTORDR5.Item("CUST_CONTACT") = String.Empty

            ' Remove trailing 0000 on a Zip Code
            If (rowSOTORDR5.Item("CUST_ZIP_CODE") & String.Empty).ToString.Length = 9 Then
                If (rowSOTORDR5.Item("CUST_ZIP_CODE") & String.Empty).ToString.Substring(5) = "0000" Then
                    rowSOTORDR5.Item("CUST_ZIP_CODE") = rowSOTORDR5.Item("CUST_ZIP_CODE").ToString.Substring(0, 5)
                End If
            End If

            rowSOTORDR5.Item("CUST_PHONE") = String.Empty
            rowSOTORDR5.Item("CUST_EXT") = String.Empty
            rowSOTORDR5.Item("CUST_FAX") = String.Empty
            rowSOTORDR5.Item("CUST_EMAIL") = String.Empty

        End Sub

        Private Sub CreateOrderShipTo(ByVal ORDR_NO As String, ByRef rowSOTORDRX As DataRow, ByVal shipToPatient As Boolean)

            Dim rowSOTORDR5 As DataRow = dst.Tables("SOTORDR5").NewRow

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

        End Sub

        Private Function CreateSalesOrder(ByVal CreateShipTo As Boolean, ByRef ORDR_NO As String) As Boolean

            Try
                CreateSalesOrder = False
                ORDR_NO = String.Empty

                ' See if we have any data to process
                If dst.Tables("SOTORDRX").Rows.Count = 0 Then
                    Return False
                End If

                'Dim ORDR_NO As String = String.Empty
                Dim CUST_CODE As String = String.Empty
                Dim CUST_SHIP_TO_NO As String = String.Empty
                Dim ORDR_QTY As Integer = 0
                Dim shipToPatient As Boolean = False
                Dim ORDR_SOURCE As String = String.Empty

                Dim ITEM_CODE As String = String.Empty

                Dim sql As String = String.Empty
                Dim sqlSalesOrder As String = String.Empty

                Dim SOTORDR1ErrorCodes As String = String.Empty
                Dim SOTORDR2ErrorCodes As String = String.Empty

                Dim rowSOTORDR1 As DataRow = Nothing
                Dim rowSOTORDR2 As DataRow = Nothing
                Dim rowSOTORDR3 As DataRow = Nothing

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

                CUST_CODE = ABSolution.ASCMAIN1.Format_Field(CUST_CODE, "CUST_CODE")
                If CUST_SHIP_TO_NO.Length > 0 Then
                    CUST_SHIP_TO_NO = ABSolution.ASCMAIN1.Format_Field(CUST_SHIP_TO_NO, "CUST_SHIP_TO_NO")
                End If

                rowARTCUST1 = baseClass.LookUp("ARTCUST1", CUST_CODE)
                rowARTCUST2 = baseClass.LookUp("ARTCUST2", New String() {CUST_CODE, CUST_SHIP_TO_NO})
                rowARTCUST3 = baseClass.LookUp("ARTCUST3", CUST_CODE)

                If CreateShipTo = True AndAlso CUST_SHIP_TO_NO.Length > 0 AndAlso rowARTCUST2 Is Nothing Then
                    CreateCustomerShipTo(CUST_CODE, CUST_SHIP_TO_NO, rowSOTORDRX)
                End If

                rowSOTORDR1 = dst.Tables("SOTORDR1").NewRow
                rowSOTORDR1.Item("ORDR_NO") = ORDR_NO
                rowSOTORDR1.Item("ORDR_TYPE_CODE") = rowSOTORDRX.Item("ORDR_TYPE_CODE") & String.Empty
                rowSOTORDR1.Item("CUST_CODE") = CUST_CODE
                rowSOTORDR1.Item("CUST_SHIP_TO_NO") = CUST_SHIP_TO_NO
                rowSOTORDR1.Item("ORDR_STATUS") = "O"
                rowSOTORDR1.Item("ORDR_LOCK_SHIP_VIA") = (rowSOTORDRX.Item("ORDR_LOCK_SHIP_VIA") & String.Empty).ToString.Trim
                dst.Tables("SOTORDR1").Rows.Add(rowSOTORDR1)

                ORDR_SOURCE = (rowSOTORDRX.Item("ORDR_SOURCE") & String.Empty).ToString.Trim
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
                                                                     "VV", New String() {ORDR_SOURCE, (rowSOTORDRX.Item("SHIP_VIA_CODE") & String.Empty).ToString.ToUpper.Trim})
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

                Me.SetOrderCustomerInfo(CUST_CODE, CUST_SHIP_TO_NO, rowSOTORDR1, errorCodes)
                Me.AddCharNoDups(errorCodes, SOTORDR1ErrorCodes)

                If rowSOTORDR1.Item("SHIP_VIA_CODE") & String.Empty = String.Empty Then
                    Me.AddCharNoDups("V", SOTORDR1ErrorCodes)
                End If

                ' Ship To
                If CUST_SHIP_TO_NO.Length > 0 Then
                    Me.SetShipToAttributes(rowSOTORDR1, SOTORDR1ErrorCodes)
                End If

                ' Set DPD settings
                If shipToPatient = True Then
                    Me.SetDPDShipViaSettings(rowSOTORDR1, SOTORDR1ErrorCodes)
                End If

                ' ************ Other Order Header Fields ************
                If IsDate(rowSOTORDRX.Item("ORDR_DATE") & String.Empty) Then
                    rowSOTORDR1.Item("ORDR_DATE") = CDate(rowSOTORDRX.Item("ORDR_DATE") & String.Empty).ToString("dd-MMM-yyyy")
                Else
                    rowSOTORDR1.Item("ORDR_DATE") = DateTime.Now.ToString("dd-MMM-yyyy")
                End If

                rowSOTORDR1.Item("ORDR_CUST_PO") = String.Empty
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

                Select Case rowSOTORDR1.Item("ORDR_SOURCE") & String.Empty
                    Case "X" ' XML

                    Case "O" ' OptiPort

                    Case "V" ' Vision Web

                    Case "F" ' eyeFinity

                    Case "C" ' Customer Excel
                        rowSOTORDR1.Item("ORDR_CALLER_NAME") = "Excel Order"
                    Case "Y" ' Eyeconic
                        rowSOTORDR1.Item("ORDR_CALLER_NAME") = "eyeconic.com"
                End Select

                Dim ORDR_COMMENT As String = (rowSOTORDRX.Item("ORDR_COMMENT") & String.Empty).ToString.Trim
                If ORDR_COMMENT.Length > 0 Then
                    Me.AddCharNoDups("R", SOTORDR1ErrorCodes)

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
                        Me.AddCharNoDups("Q", SOTORDR1ErrorCodes)
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

                rowSOTORDR1.Item("ORDR_REL_HOLD_CODES") = SOTORDR1ErrorCodes.Trim

                CreateOrderBillTo(ORDR_NO)
                CreateOrderShipTo(ORDR_NO, rowSOTORDRX, shipToPatient)

                ' May not be needed the XST tables have the ORDR_NO and ORDR_LNO values
                'If shipToPatient Then
                '    CreateOrderPatientBillTo(ORDR_NO, rowSOTORDRX)
                'End If

                CreateSalesOrderTax(ORDR_NO)

                If (rowSOTORDR1.Item("ORDR_REL_HOLD_CODES") & String.Empty).ToString.Trim.Length = 0 Then
                    Me.GetSalesOrderUnitPrices(ORDR_NO)
                    Me.UpdateSalesDollars(ORDR_NO)
                Else
                    rowSOTORDR1.Item("ORDR_STATUS_WEB") = ORDR_SOURCE
                End If

                rowSOTORDR1.Item("INIT_DATE") = DateTime.Now
                rowSOTORDR1.Item("LAST_DATE") = DateTime.Now

                rowSOTORDR1.Item("INIT_OPER") = ABSolution.ASCMAIN1.USER_ID
                rowSOTORDR1.Item("LAST_OPER") = ABSolution.ASCMAIN1.USER_ID

                CreateSalesOrder = True

            Catch ex As Exception
                CreateSalesOrder = False
                RecordLogEntry("CreateSalesOrder: " & ex.Message)
            End Try

        End Function

        Private Sub CreateSalesOrderTax(ByVal ORDR_NO As String)

            Dim CUST_SHIP_TO_ZIP_TAX As String = String.Empty
            Dim CUST_SHIP_TO_STATE As String = String.Empty
            Dim STAX_EXEMPT As String = String.Empty
            Dim STAX_CODE As String = String.Empty
            Dim STAX_RATE As Double = 0

            Dim rowSOTORDR1 As DataRow = dst.Tables("SOTORDR1").Select("ORDR_NO = '" & ORDR_NO & "'")(0)

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

        End Sub

        Private Function GetOrderSalesTaxByState(ByVal rowSOTORDR1 As DataRow, ByVal tblSOTORDR2 As DataTable) As Double

            Dim ORDR_NO As String = rowSOTORDR1.Item("ORDR_NO") & String.Empty
            Dim CUST_CODE As String = rowSOTORDR1.Item("CUST_CODE") & String.Empty
            Dim CUST_SHIP_TO_NO As String = rowSOTORDR1.Item("CUST_SHIP_TO_NO") & String.Empty

            CreateSalesOrderTax(ORDR_NO)

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

            Return Math.Round((taxableAmount * rowSOTORDR1.Item("STAX_RATE")) / 100, 2, MidpointRounding.AwayFromZero)

        End Function

        Private Sub GetSalesOrderUnitPrices(ByVal ORDR_NO As String)

            Dim rowSOTORDR1 As DataRow = dst.Tables("SOTORDR1").Select("ORDR_NO = '" & ORDR_NO & "'")(0)

            Me.TestAuthorizationsAndBlocks(rowSOTORDR1)

            clsSOCORDR1.AffiliateFreeShipping()
            clsSOCORDR1.Price_and_Qty()

            ' Added on 1/22/2009 as per walter
            If clsSOCORDR1.SHIP_VIA_CODE_switch_to.Trim.Length > 0 Then
                If (dst.Tables("SOTORDR1").Rows(0).Item("ORDR_LOCK_SHIP_VIA") & String.Empty) <> "1" Then
                    dst.Tables("SOTORDR1").Rows(0).Item("SHIP_VIA_CODE") = clsSOCORDR1.SHIP_VIA_CODE_switch_to.Trim
                End If
            End If

            If clsSOCORDR1.ORDR_NO_FREIGHT & String.Empty = "1" _
                AndAlso dst.Tables("SOTORDR1").Rows(0).Item("ORDR_NO_FREIGHT") & String.Empty <> "1" Then
                dst.Tables("SOTORDR1").Rows(0).Item("ORDR_NO_FREIGHT") = "1"
                dst.Tables("SOTORDR1").Rows(0).Item("REASON_CODE_NO_FRT") = clsSOCORDR1.REASON_CODE_NO_FRT
            End If

        End Sub

        ''' <summary>
        ''' Overrides Order Settings for DPD orders
        ''' </summary>
        ''' <param name="rowSOTORDR1"></param>
        ''' <param name="errorCodes"></param>
        ''' <remarks></remarks>
        Private Sub SetDPDShipViaSettings(ByRef rowSOTORDR1 As DataRow, ByRef errorCodes As String)

            If (rowSOTORDR1.Item("ORDR_DPD") & String.Empty) <> "1" Then Exit Sub
            If (rowSOTORDR1.Item("ORDR_LOCK_SHIP_VIA") & String.Empty) = "1" Then Exit Sub

            Dim SHIP_VIA_CODE_DPD As String = String.Empty

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
                SHIP_VIA_CODE_DPD = DpdDefaultShipViaCode
            End If

            rowSOTSVIA1 = baseClass.LookUp("SOTSVIA1", SHIP_VIA_CODE_DPD)

            If rowSOTSVIA1 IsNot Nothing Then
                rowSOTORDR1.Item("SHIP_VIA_CODE") = rowSOTSVIA1.Item("SHIP_VIA_CODE") & String.Empty
                errorCodes = Replace(errorCodes, "V", String.Empty)
            End If

        End Sub

        ''' <summary>
        ''' Record sales order event
        ''' </summary>
        ''' <param name="ORDR_NO"></param>
        ''' <param name="EVENT_DESC"></param>
        ''' <remarks></remarks>
        Private Sub Record_Event(ByVal ORDR_NO As String, ByVal EVENT_DESC As String)
            Dim row As DataRow = dst.Tables("SOTORDRE").NewRow
            row.Item("ORDR_NO") = ORDR_NO
            row.Item("INIT_DATE") = DateTime.Now
            row.Item("INIT_OPER") = ABSolution.ASCMAIN1.USER_ID
            row.Item("EVENT_DESC") = EVENT_DESC
            dst.Tables("SOTORDRE").Rows.Add(row)
        End Sub

        ''' <summary>
        ''' Sets Order Header Information based on the Bill To Customer Attributes
        ''' </summary>
        ''' <param name="CUST_CODE"></param>
        ''' <param name="CUST_SHIP_TO_NO"></param>
        ''' <param name="rowSOTORDR1"></param>
        ''' <param name="errorCodes"></param>
        ''' <remarks></remarks>
        Private Sub SetOrderCustomerInfo(ByVal CUST_CODE As String, ByVal CUST_SHIP_TO_NO As String, ByRef rowSOTORDR1 As DataRow, ByRef errorCodes As String)

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
                rowSOTORDR1.Item("SHIP_VIA_CODE") = String.Empty
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

                'CreateSalesOrderTax(ORDR_NO)

                TERM_CODE = rowARTCUST1.Item("TERM_CODE") & String.Empty
                rowTATTERM1 = baseClass.LookUp("TATTERM1", TERM_CODE)
                If rowTATTERM1 IsNot Nothing Then
                    rowSOTORDR1.Item("TERM_CODE") = TERM_CODE
                Else
                    AddCharNoDups("T", errorCodes)
                End If

                rowSOTORDR1.Item("POST_CODE") = rowARTCUST1.Item("POST_CODE") & String.Empty
                rowSOTORDR1.Item("SREP_CODE") = rowARTCUST1.Item("SREP_CODE") & String.Empty
                rowSOTORDR1.Item("ORDR_SHIP_COMPLETE") = rowARTCUST1.Item("CUST_SHIP_COMPLETE") & String.Empty
                rowSOTORDR1.Item("ORDR_NO_SAMPLE_SURCHARGE") = rowARTCUST1.Item("NO_SAMPLE_SURCHARGE") & String.Empty
                rowSOTORDR1.Item("ORDR_NO_SAMPLE_HANDLING_FEE") = rowARTCUST1.Item("NO_SAMPLE_HANDLING_FEE") & String.Empty
            Else
                rowSOTORDR1.Item("CUST_NAME") = "Unknown Customer"
                rowSOTORDR1.Item("CUST_BILL_TO_CUST") = String.Empty
                Me.AddCharNoDups("B", errorCodes)
            End If

            If rowSOTORDR1.Item("ORDR_DPD") & String.Empty = "1" Then
                SetDPDShipViaSettings(rowSOTORDR1, errorCodes)
            ElseIf (rowSOTORDR1.Item("ORDR_LOCK_SHIP_VIA") & String.Empty) <> "1" Then
                If rowARTCUST2 IsNot Nothing AndAlso (rowARTCUST2.Item("CUST_SHIP_TO_SHIP_VIA_CODE") & String.Empty <> String.Empty) Then
                    rowSOTORDR1.Item("SHIP_VIA_CODE") = rowARTCUST2.Item("CUST_SHIP_TO_SHIP_VIA_CODE") & String.Empty
                    rowSOTSVIA1 = baseClass.LookUp("SOTSVIA1", rowSOTORDR1.Item("SHIP_VIA_CODE"))
                    If rowSOTSVIA1 IsNot Nothing Then
                        rowSOTORDR1.Item("SHIP_VIA_DESC") = rowSOTSVIA1.Item("SHIP_VIA_DESC") & String.Empty
                    End If
                ElseIf rowARTCUST3 IsNot Nothing AndAlso (rowSOTORDR1.Item("SHIP_VIA_CODE") & String.Empty = String.Empty) Then
                    If rowSOTORDR1.Item("ORDR_DPD") & String.Empty = "1" Then
                        rowSOTORDR1.Item("SHIP_VIA_CODE") = rowARTCUST3.Item("SHIP_VIA_CODE_DPD") & String.Empty
                        If rowSOTORDR1.Item("SHIP_VIA_CODE") & String.Empty = String.Empty Then
                            rowSOTORDR1.Item("SHIP_VIA_CODE") = DpdDefaultShipViaCode
                        End If
                    Else
                        rowSOTORDR1.Item("SHIP_VIA_CODE") = rowARTCUST3.Item("SHIP_VIA_CODE") & String.Empty
                    End If

                    rowSOTSVIA1 = baseClass.LookUp("SOTSVIA1", rowSOTORDR1.Item("SHIP_VIA_CODE"))
                    If rowSOTSVIA1 IsNot Nothing Then
                        rowSOTORDR1.Item("SHIP_VIA_DESC") = rowSOTSVIA1.Item("SHIP_VIA_DESC") & String.Empty
                    End If
                End If
            End If

            If rowARTCUST3 IsNot Nothing Then
                rowSOTORDR1.Item("FRT_CONT_NO") = rowARTCUST3.Item("FRT_CONT_NO") & String.Empty
            End If
        End Sub

        ''' <summary>
        ''' Sets Order Detail Information based on the Item's Attributes
        ''' </summary>
        ''' <param name="rowSOTORDR2"></param>
        ''' <param name="errorCodes"></param>
        ''' <remarks></remarks>
        Private Sub SetItemInfo(ByRef rowSOTORDR2 As DataRow, ByRef errorCodes As String)

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
                Me.AddCharNoDups("I", errorCodes)
            Else
                rowSOTORDR2.Item("ITEM_CODE") = rowICTITEM1.Item("ITEM_CODE") & String.Empty
                rowSOTORDR2.Item("ITEM_DESC") = rowICTITEM1.Item("ITEM_DESC") & String.Empty
                rowSOTORDR2.Item("ITEM_DESC2") = rowICTITEM1.Item("ITEM_DESC2") & String.Empty
                rowSOTORDR2.Item("ITEM_UOM") = "EA"
                rowSOTORDR2.Item("PRICE_CATGY_CODE") = rowICTITEM1.Item("PRICE_CATGY_CODE") & String.Empty

                If rowICTITEM1.Item("ITEM_ORDER_CODE") & String.Empty = "X" OrElse rowICTITEM1.Item("ITEM_STATUS") & String.Empty = "I" Then
                    Me.AddCharNoDups("F", errorCodes)
                End If
            End If

        End Sub

        ''' <summary>
        ''' Ship Tos have Overrides to the Customer Master data.
        ''' </summary>
        ''' <param name="rowSOTORDR1"></param>
        ''' <param name="errorCodes"></param>
        ''' <remarks></remarks>
        Private Sub SetShipToAttributes(ByRef rowSOTORDR1 As DataRow, ByRef errorCodes As String)

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

                If (rowSOTORDR1.Item("ORDR_LOCK_SHIP_VIA") & String.Empty) <> "1" Then
                    If rowSOTORDR1.Item("ORDR_DPD") & String.Empty = "1" Then
                        'Done Elsewhere
                        'SetDPDShipViaSettings(rowSOTORDR1, errorCodes)
                    ElseIf (rowARTCUST2.Item("CUST_SHIP_TO_SHIP_VIA_CODE") & String.Empty).ToString.Length > 0 Then
                        rowSOTORDR1.Item("SHIP_VIA_CODE") = rowARTCUST2.Item("CUST_SHIP_TO_SHIP_VIA_CODE") & String.Empty
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
                Me.AddCharNoDups("S", errorCodes)
            End If
        End Sub

        Private Sub TestAuthorizationsAndBlocks(ByRef rowSOTORDR1 As DataRow)

            If clsSOCORDR1 Is Nothing Then Exit Sub

            Dim CUST_CODE As String = (rowSOTORDR1.Item("CUST_CODE") & String.Empty).ToString.Trim
            Dim CUST_SHIP_TO_NO As String = (rowSOTORDR1.Item("CUST_SHIP_TO_NO") & String.Empty).ToString.Trim
            Dim ITEM_LIST As String = String.Empty
            Dim ORDR_NO As String = rowSOTORDR1.Item("ORDR_NO")
            Dim ORDR_REL_HOLD_CODES As String = String.Empty

            ' Remove Header Error Code
            ORDR_REL_HOLD_CODES = rowSOTORDR1.Item("ORDR_REL_HOLD_CODES") & String.Empty
            ORDR_REL_HOLD_CODES = ORDR_REL_HOLD_CODES.Replace("0", "") & String.Empty
            rowSOTORDR1.Item("ORDR_REL_HOLD_CODES") = ORDR_REL_HOLD_CODES

            ORDR_REL_HOLD_CODES = String.Empty

            For Each rowSOTORDR2 As DataRow In dst.Tables("SOTORDR2").Select("ISNULL(ORDR_REL_HOLD_CODES,'@@@') <> '@@@'")
                ITEM_LIST = "'" & rowSOTORDR2.Item("ITEM_CODE") & "'"
                ORDR_REL_HOLD_CODES = rowSOTORDR2.Item("ORDR_REL_HOLD_CODES") & String.Empty

                ' Remove Detail Error Code
                ORDR_REL_HOLD_CODES = ORDR_REL_HOLD_CODES.Replace("0", "") & String.Empty

                Dim errors As String = clsSOCORDR1.TestAuthorizationsAndBlocks(CUST_CODE, CUST_SHIP_TO_NO, ITEM_LIST, False)

                errors = errors.Trim
                If errors.Length = 0 Then Continue For

                ORDR_REL_HOLD_CODES &= "0"

                For Each authError As String In errors.Split(vbCr)
                    authError = authError.Trim
                    If authError.Length > 0 Then
                        CreateOrderErrorRecord(ORDR_NO, 0, "0", authError)
                    End If
                Next

                ' Place the error on the header
                If Not (rowSOTORDR1.Item("ORDR_REL_HOLD_CODES") & String.Empty).ToString.Contains("0") Then
                    rowSOTORDR1.Item("ORDR_REL_HOLD_CODES") &= "0"
                End If

                rowSOTORDR2.Item("ORDR_REL_HOLD_CODES") = ORDR_REL_HOLD_CODES
            Next
        End Sub

        ''' <summary>
        ''' Truncates a fields value if the length is longer the the max length of the field
        ''' </summary>
        ''' <param name="fieldValue"></param>
        ''' <param name="TableName"></param>
        ''' <param name="FieldName"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Private Function TruncateField(ByVal fieldValue As String, ByVal TableName As String, ByVal FieldName As String) As String

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

            Return rValue
        End Function

        ''' <summary>
        ''' Updates the Freight for an Order
        ''' </summary>
        ''' <param name="OrderNumber"></param>
        ''' <remarks></remarks>
        Private Sub UpdateSalesDollars(ByVal OrderNumber As String)

            Try
                Dim rowSOTORDR1 As DataRow = dst.Tables("SOTORDR1").Rows(0)
                Dim tblSOTORDR2 As DataTable = dst.Tables("SOTORDR2")
                Dim ORDR_FREIGHT As Double = 0

                Dim ORDR_TOTAL_QTY As Integer = Val(dst.Tables("SOTORDR2").Compute("SUM(ORDR_QTY)", "ORDR_NO = '" & OrderNumber & "'") & String.Empty)
                Dim ORDR_TOTAL_SALES As Double = Val(dst.Tables("SOTORDR2").Compute("SUM(ORDR_LNO_EXT)", "ORDR_NO = '" & OrderNumber & "'") & String.Empty)
                Dim shipToPatient As Boolean = rowSOTORDR1.Item("ORDR_DPD") & String.Empty = "1"

                rowSOTORDR1.Item("ORDR_SALES") = Math.Round(ORDR_TOTAL_SALES, 2, MidpointRounding.AwayFromZero)

                If rowSOTORDR1.Item("ORDR_NO_FREIGHT") & String.Empty <> "1" Then
                    ORDR_FREIGHT = TAC.SOCMAIN1.Get_INV_FREIGHT(baseClass, rowSOTORDR1.Item("CUST_CODE"), _
                        rowSOTORDR1.Item("CUST_SHIP_TO_NO") & "", _
                        rowSOTORDR1.Item("SHIP_VIA_CODE"), rowSOTORDR1.Item("ORDR_DATE"), ORDR_TOTAL_QTY, ORDR_TOTAL_SALES, shipToPatient, "E")
                End If

                rowSOTORDR1.Item("ORDR_FREIGHT") = Math.Round(ORDR_FREIGHT, 2, MidpointRounding.AwayFromZero)

                rowSOTORDR1.Item("ORDR_STAX") = Me.GetOrderSalesTaxByState(rowSOTORDR1, dst.Tables("SOTORDR2"))
                rowSOTORDR1.Item("ORDR_TOTAL_AMT") = Val(rowSOTORDR1.Item("ORDR_SALES") & String.Empty) _
                    + Val(rowSOTORDR1.Item("ORDR_FREIGHT") & String.Empty) _
                    + Val(rowSOTORDR1.Item("ORDR_STAX") & String.Empty) _
                    + Val(rowSOTORDR1.Item("ORDR_SAMPLE_SURCHARGE") & String.Empty) _
                    + Val(rowSOTORDR1.Item("ORDR_MISC_CHG_AMT") & String.Empty)

            Catch ex As Exception
                RecordLogEntry(ex.Message)
            End Try
        End Sub

#End Region

#Region "DataSet Functions"

        Private Sub ClearDataSetTables(ByVal ClearXMTtables As Boolean)

            With dst
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

        End Sub

        Private Sub Dependent_Updates(ByVal ORDR_NO As String, ByVal S As Integer)

            Dim PLUS_OR_MINUS As String = "+1*"
            Dim sql As String = String.Empty

            If S = -1 Then
                PLUS_OR_MINUS = "-1*"
            End If

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

        End Sub

        Private Sub LoadTablesForPricing()

            With dst

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

                clsSOCORDR1 = New TAC.SOCORDR1(SOTINVH2_PC, SOTORDRP, SOTORDR2_pricing, baseClass.clsASCBASE1)

            End With

        End Sub

        Private Sub PrepareDatasetEntries()

            Dim sql As String = String.Empty

            dst = baseClass.clsASCBASE1.dst

            With dst

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

            End With

        End Sub

        Private Sub UpdateDataSetTables()

            Dim sql As String = String.Empty

            With baseClass
                Try
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

                Catch ex As Exception
                    .Rollback()
                    RecordLogEntry("UpdateDataSetTables  : " & ex.Message)
                End Try
            End With

        End Sub

#End Region

#Region "Log Procedures"

        Private Sub OpenLogFile()
            logFilename = Format(Now, "yyyyMMddHHmm") & ".log"
            If logStreamWriter IsNot Nothing Then
                logStreamWriter.Close()
                logStreamWriter.Dispose()
            End If
            logStreamWriter = New System.IO.StreamWriter(filefolder & logFilename, True)
        End Sub

        Private Sub RecordLogEntry(ByVal message As String)
            logStreamWriter.WriteLine(DateTime.Now & ": " & message)
        End Sub

        Public Sub CloseLog()
            If logStreamWriter IsNot Nothing Then
                logStreamWriter.Close()
                logStreamWriter.Dispose()
            End If
        End Sub

#End Region

    End Class

End Namespace


