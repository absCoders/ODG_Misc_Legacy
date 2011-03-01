﻿Imports InvoiceEmail.Extensions

Namespace InoviceEmail

    Public Class InvoiceEmailer

        Private WithEvents importTimer As System.Threading.Timer
        Declare Function ProcessIdToSessionId Lib "kernel32.dll" (ByVal dwProcessId As Int32, ByRef pSessionId As Int32) As Int32

#Region "Service Variables"

        Private baseClass As ABSolution.ASFBASE1

        Private emailInProcess As Boolean = False
        Private logFilename As String = String.Empty
        Private filefolder As String = String.Empty
        Private logStreamWriter As System.IO.StreamWriter
        Private dst As DataSet

        Private Const testMode As Boolean = False

        Private sqlDPD As String = String.Empty
        Private sqlCRM As String = String.Empty
        Private sqlECP As String = String.Empty

#End Region

#Region "Instaniate Service"

        Public Sub New()

        End Sub

#End Region

#Region "Data Management"

        Private Sub MainProcess()
            Try

                ' Prevent the code from firing if still importing
                If emailInProcess Then Exit Sub
                emailInProcess = True

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
                        EmailInvoicesToCustomers()
                    End If
                End If

                If testMode Then RecordLogEntry("Exit MainProcess.")
                RecordLogEntry("Closing Log file.")
                CloseLog()

            Catch ex As Exception
                RecordLogEntry("MainProcess: " & ex.Message)
            Finally
                emailInProcess = False
            End Try

        End Sub

        Public Sub LogIn()
            importTimer = New System.Threading.Timer _
                (New System.Threading.TimerCallback(AddressOf MainProcess), Nothing, 3000, 3300000) ' every 55 mins 
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
                logFilename = String.Empty
                filefolder = String.Empty

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

                Dim folder_prefix As String
                'MsgBox("2a")
                If UCase(My.Application.Info.DirectoryPath) Like "C:\VS\*" Then
                    ABSolution.ASCMAIN1.Running_in_VS = True
                    folder_prefix = "\..\..\..\..\"
                    ABSolution.ASCMAIN1.CLIENT_CODE = UCase(Mid(My.Application.Info.DirectoryPath, 7, 3))
                Else
                    ABSolution.ASCMAIN1.Running_in_VS = False
                    folder_prefix = "\..\"
                    ABSolution.ASCMAIN1.CLIENT_CODE = UCase(Split(My.Application.Info.DirectoryPath, "\")(3))
                End If
                'MsgBox("2b")

                ABSolution.ASCMAIN1.Folders.Add("Images", ABSolution.ASCMAIN1.GetPath(My.Application.Info.DirectoryPath & folder_prefix & "Images\"))
                ABSolution.ASCMAIN1.Folders.Add("Reports", ABSolution.ASCMAIN1.GetPath(My.Application.Info.DirectoryPath & folder_prefix & "Reports\"))
                ABSolution.ASCMAIN1.Folders.Add("DataSets", ABSolution.ASCMAIN1.GetPath(My.Application.Info.DirectoryPath & folder_prefix & "DataSets\"))
                ABSolution.ASCMAIN1.Folders.Add("Temp", ABSolution.ASCMAIN1.GetPath(My.Application.Info.DirectoryPath & folder_prefix & "Temp\"))
                ABSolution.ASCMAIN1.Folders.Add("Work", ABSolution.ASCMAIN1.GetPath(My.Application.Info.DirectoryPath & folder_prefix & "Work\"))
                ABSolution.ASCMAIN1.Folders.Add("bin", ABSolution.ASCMAIN1.GetPath(My.Application.Info.DirectoryPath & folder_prefix & "bin\"))
                ABSolution.ASCMAIN1.Folders.Add("Help", ABSolution.ASCMAIN1.GetPath(My.Application.Info.DirectoryPath & folder_prefix & "Help\"))
                ABSolution.ASCMAIN1.Folders.Add("Archive", ABSolution.ASCMAIN1.GetPath(My.Application.Info.DirectoryPath & folder_prefix & "Archive\"))
                ABSolution.ASCMAIN1.Folders.Add("Attach", ABSolution.ASCMAIN1.GetPath(My.Application.Info.DirectoryPath & folder_prefix & "Attach\"))
                ABSolution.ASCMAIN1.Folders.Add("root", ABSolution.ASCMAIN1.GetPath(My.Application.Info.DirectoryPath & folder_prefix))

                If My.Computer.Name = "WJZLAP" Then
                    ABSolution.ASCMAIN1.Folders.Add("Oracle", "C:\oracle\product\10.2.0\db_1\")
                Else
                    ABSolution.ASCMAIN1.Folders.Add("Oracle", "C:\oracle\product\10.2.0\Client_1\")
                End If

                ABSolution.ASCMAIN1.ActiveForm = baseClass


                ' Sql Statements
                sqlDPD = "Select INV_NO FROM SOTINVH1"
                sqlDPD &= " WHERE INV_DPD_PRINT_IND = 'D'"
                sqlDPD &= " AND ORDR_TYPE_CODE <> 'B2C'"
                sqlDPD &= " AND CUST_CODE = :PARM1"

                sqlECP = "Select INV_NO FROM SOTINVH1"
                sqlECP &= " WHERE INV_DPD_PRINT_IND = 'D'"
                sqlECP &= " AND ORDR_TYPE_CODE = 'B2C'"
                sqlECP &= " AND CUST_CODE_B2B = :PARM1"

                sqlCRM = "SELECT RTRN_NO INV_NO FROM SOTRTRN1 WHERE INV_PRINTED = '0' AND CUST_CODE = :PARM1"

                If testMode Then RecordLogEntry("Exit InitializeSettings.")

                Return True

            Catch ex As Exception
                RecordLogEntry("InitializeSettings: " & ex.Message)
                Return False
            End Try

        End Function

        Private Function GetSessionId() As Int32
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

        Private Function EmailInvoicesToCustomers() As Int16

            Dim numEmails As Int16 = 0
            Dim sql As String = String.Empty

            Try


                If testMode Then RecordLogEntry("Enter EmailInvoicesToCustomers.")

                Dim svcConfig As New ServiceConfig
                Dim milTime As String = svcConfig.StartEmailing
                If (milTime = "0000") Then
                    RecordLogEntry("EmailInvoicesToCustomers: Start time set 0000, indicates do not set invoices")
                    Return numEmails
                ElseIf (milTime.Length <> 4) Then
                    RecordLogEntry("EmailInvoicesToCustomers: Invalid Military time to start emailing invoices")
                    Return numEmails
                Else
                    If (CInt(milTime.Substring(0, 2)) < 12) Then
                        milTime = milTime.Substring(0, 2) + ":" + milTime.Substring(2, 2) + "AM"
                    Else
                        milTime = CStr(CInt(milTime.Substring(0, 2)) - 12) + ":" + milTime.Substring(2, 2) + "PM"
                    End If
                End If

                Dim localTime As Date = DateTime.Now.ToLocalTime
                Dim processTime As Date = CDate(DateTime.Now.ToString("MM/dd/yyyy") & " " & milTime)

                Select Case DateTime.Compare(localTime, processTime)
                    Case Is < 0
                        RecordLogEntry("EmailInvoicesToCustomers: To early to process invoices")
                        'Return numEmails
                    Case Else
                        ' No message at this time
                End Select


                ' Process email invoices by customer
                sql = "SELECT CUST_CODE, 'D' EMAIL_TYPE  FROM ARTCUST1 WHERE DPD_COPIES = 'E'"
                sql &= " Union "
                sql &= "SELECT CUST_CODE, 'C' EMAIL_TYPE  FROM ARTCUST1 WHERE CRM_COPIES = 'E'"
                sql &= " Union "
                sql &= "SELECT CUST_CODE, 'E' EMAIL_TYPE  FROM ARTCUST1 WHERE ECP_COPIES = 'E'"

                Dim tblCustomers As DataTable = ABSolution.ASCDATA1.GetDataTable(sql)

                If tblCustomers Is Nothing OrElse tblCustomers.Rows.Count = 0 Then
                    Return 0
                End If

                Dim CUST_CODE As String = String.Empty

                Dim dpdFile As String = String.Empty
                Dim crmFile As String = String.Empty
                Dim ecpFile As String = String.Empty

                Dim dpdInvoices As String = String.Empty
                Dim crmInvoices As String = String.Empty
                Dim ecpInvoices As String = String.Empty

                Dim tblInvoices As DataTable = Nothing
                Dim rowARTCUST1 As DataRow = Nothing
                Dim custEmailaddress As String = String.Empty
                Dim invoiceNumbers As String = String.Empty
                Dim attachments As String = String.Empty

                For Each rowCustomer As DataRow In ABSolution.ASCDATA1.SelectDistinct(tblCustomers, New String() {"CUST_CODE"}).Rows
                    CUST_CODE = rowCustomer.Item("CUST_CODE") & String.Empty
                    dpdFile = String.Empty
                    crmFile = String.Empty
                    ecpFile = String.Empty

                    dpdInvoices = String.Empty
                    crmInvoices = String.Empty
                    ecpInvoices = String.Empty

                    custEmailaddress = String.Empty

                    rowARTCUST1 = baseClass.LookUp("ARTCUST1", CUST_CODE)
                    If rowARTCUST1 Is Nothing Then
                        RecordLogEntry("Customer not found - " & CUST_CODE)
                        Continue For
                    ElseIf (rowARTCUST1.Item("CUST_EMAIL") & String.Empty).ToString.Trim.Length = 0 Then
                        RecordLogEntry("Customer does not have an email address: " & CUST_CODE)
                        Continue For
                    End If

                    custEmailaddress = (rowARTCUST1.Item("CUST_EMAIL") & String.Empty).ToString.Trim

                    For Each rowExport As DataRow In tblCustomers.Select("CUST_CODE = '" & CUST_CODE & "'", "EMAIL_TYPE")
                        invoiceNumbers = String.Empty

                        Select Case rowExport.Item("EMAIL_TYPE")
                            Case "D"
                                sql = sqlDPD
                            Case "C"
                                sql = sqlCRM
                            Case "E"
                                sql = sqlECP
                        End Select

                        For Each rowInvoices As DataRow In ABSolution.ASCDATA1.GetDataTable(sql, "", "V", New Object() {CUST_CODE}).Rows
                            invoiceNumbers &= ", '" & rowInvoices.Item("INV_NO") & "'"
                        Next

                        If invoiceNumbers.Length = 0 Then Continue For
                        invoiceNumbers = invoiceNumbers.Substring(1).Trim

                        Select Case rowExport.Item("EMAIL_TYPE")
                            Case "D"
                                dpdInvoices = invoiceNumbers
                                dpdFile = CreateDPDInvoiceFile(invoiceNumbers, CUST_CODE)
                            Case "C"
                                crmInvoices = invoiceNumbers
                                crmFile = CreateCrmFile(invoiceNumbers, CUST_CODE)
                            Case "E"
                                ecpInvoices = invoiceNumbers
                                ecpFile = CreateECPInvoiceFile(invoiceNumbers, CUST_CODE)
                        End Select
                    Next

                    attachments = String.Empty
                    If dpdFile.Length > 0 Then attachments &= ";" & dpdFile
                    If crmFile.Length > 0 Then attachments &= ";" & crmFile
                    If ecpFile.Length > 0 Then attachments &= ";" & ecpFile

                    If attachments.Length = 0 Then Continue For

                    attachments = attachments.Substring(1)

                    EmailDocument(custEmailaddress, "odg@opticaldg.com", "ODG Invoices and Credit Memos", attachments)

                    UpdateDataSetTables(dpdInvoices, "D")
                    UpdateDataSetTables(crmInvoices, "C")
                    UpdateDataSetTables(ecpInvoices, "E")

                    For Each file As String In attachments.Split(";")
                        file = file.Trim
                        If file.Length = 0 Then Continue For
                        If My.Computer.FileSystem.FileExists(file) Then
                            My.Computer.FileSystem.DeleteFile(file)
                        End If
                    Next

                Next

                If testMode Then RecordLogEntry("Exit InitializeSettings.")

                Return numEmails

            Catch ex As Exception
                RecordLogEntry("EmailInvoicesToCustomers: " & ex.Message)

            Finally
                RecordLogEntry(numEmails & " Emails sent out")
            End Try

        End Function

        Private Function CreateDPDInvoiceFile(ByVal InvoiceNumbers As String, ByVal CustomerCode As String) As String

            Try

                Dim reportNo As String = String.Empty
                Dim outputFilenames As String = String.Empty
                Dim generatedReport As String = String.Empty

                Dim rptSORINVC1 As ABSolution.ASFSRPTM
                rptSORINVC1 = baseClass.Load_rptClass("SORINVC1")
                rptSORINVC1.Prepare_dst(False, "")
                rptSORINVC1.Fill_Records_RPT(InvoiceNumbers)

                With rptSORINVC1.clsASCBASE1
                    .Print_Report_Begin()
                    generatedReport = CustomerCode & "_dpd4"
                    reportNo = .Generate_Report("SORINVC4", "DPD Patient Copy", "", False, False, "", "PDF", generatedReport, False)
                    generatedReport = .F.REPORT_FILENAMES(reportNo)
                    .Print_Report_End(, True)
                End With
                outputFilenames &= ";" & generatedReport

                'Me.SetupDoctorReturnAddress()
                With rptSORINVC1.clsASCBASE1
                    .Print_Report_Begin()
                    generatedReport = CustomerCode & "_dpd1"
                    .CR_params.Add("PRE_PRINTED_FORM", "0")
                    reportNo = .Generate_Report("SORINVC1", "Dr. Copy", "", False, False, "", "PDF", generatedReport, False)
                    generatedReport = .F.REPORT_FILENAMES(reportNo)
                    .Print_Report_End(, True)
                End With
                outputFilenames &= ";" & generatedReport

                rptSORINVC1.Dispose()
                rptSORINVC1 = Nothing

                Return outputFilenames

            Catch ex As Exception
                Return String.Empty
            End Try

        End Function

        Private Function CreateECPInvoiceFile(ByVal InvoiceNumbers As String, ByVal CustomerCode As String) As String

            Try
                Dim outputFilenames As String = String.Empty
                Dim reportNo As String = String.Empty
                Dim generatedReport As String = String.Empty

                Dim rptSORINVC1 As ABSolution.ASFSRPTM
                rptSORINVC1 = baseClass.Load_rptClass("SORINVC1")
                rptSORINVC1.Prepare_dst(False, "")
                rptSORINVC1.Fill_Records_RPT(InvoiceNumbers)

                With rptSORINVC1.clsASCBASE1
                    .Print_Report_Begin()
                    generatedReport = CustomerCode & "_ecp3"
                    reportNo = .Generate_Report("SORINVC3", "B2C Patient Copy", "", False, False, "", "PDF", generatedReport, False)
                    generatedReport = .F.REPORT_FILENAMES(reportNo)
                    .Print_Report_End(, True)
                End With
                outputFilenames &= ";" & generatedReport

                'Me.SetupDoctorReturnAddress()
                With rptSORINVC1.clsASCBASE1
                    .Print_Report_Begin()
                    generatedReport = CustomerCode & "_ecp1"
                    .CR_params.Add("PRE_PRINTED_FORM", "0")
                    reportNo = .Generate_Report("SORINVC1", "Dr. Copy", "", False, False, "", "PDF", generatedReport, False)
                    generatedReport = .F.REPORT_FILENAMES(reportNo)
                    .Print_Report_End(, True)
                End With
                outputFilenames &= ";" & generatedReport

                rptSORINVC1 = New ABSolution.ASFSRPTM
                rptSORINVC1.Dispose()
                rptSORINVC1 = Nothing

                Return outputFilenames

            Catch ex As Exception
                Return String.Empty
            End Try
        End Function

        Private Function CreateCrmFile(ByVal InvoiceNumbers As String, ByVal CustomerCode As String) As String

            Try
                Dim outputFilenames As String = String.Empty
                Dim reportNo As String = String.Empty
                Dim generatedReport As String = String.Empty

                Dim rptSORRTRN1 As ABSolution.ASFSRPTM
                rptSORRTRN1 = baseClass.Load_rptClass("SORRTRN1")
                rptSORRTRN1.Prepare_dst(False, " RTRN_NO IN (" & InvoiceNumbers & ")")
                rptSORRTRN1.Fill_Records_RPT()

                With rptSORRTRN1.clsASCBASE1
                    .Print_Report_Begin()
                    generatedReport = CustomerCode & "_rtrn1"
                    .CR_params.Add("PRE_PRINTED_FORM", "0")
                    reportNo = .Generate_Report("SORRTRN1", "Credits", "", False, False, "", "PDF", generatedReport, False)
                    generatedReport = .F.REPORT_FILENAMES(reportNo)
                    .Print_Report_End(, True)
                End With
                outputFilenames &= ";" & generatedReport

                rptSORRTRN1 = New ABSolution.ASFSRPTM
                rptSORRTRN1.Dispose()
                rptSORRTRN1 = Nothing

                Return outputFilenames

            Catch ex As Exception
                Return String.Empty
            End Try

        End Function

        ''' <summary>
        ''' Sends an email using the Components created frm teh last call to CreateComponents
        ''' </summary>
        ''' <remarks></remarks>
        Private Sub EmailDocument(ByVal emailTo As String, ByVal emailFrom As String, ByVal emailSubjectText As String, ByVal attachments As String)

            If emailTo.Length = 0 OrElse emailFrom.Length = 0 Then
                Exit Sub
            End If

            Dim SEND_FROM_SIGNATURE As String = String.Empty
            Dim EMAIL_LOGO As String = String.Empty
            Dim emailCC As String = String.Empty
            Dim emailBCC As String = String.Empty
            Dim documentText As String = String.Empty

            Try

                Dim mail As New Net.Mail.MailMessage()
                mail.From = New Net.Mail.MailAddress(emailFrom, "")

                For Each sendTo As String In emailTo.Split(";")
                    If sendTo.Length > 0 Then
                        mail.To.Add(New Net.Mail.MailAddress(sendTo, ""))
                    End If
                Next

                For Each cc As String In emailCC.Split(";")
                    If cc.Length > 0 Then
                        mail.CC.Add(New Net.Mail.MailAddress(cc, ""))
                    End If
                Next

                For Each bcc As String In emailBCC.Split(";")
                    If bcc.Length > 0 Then
                        mail.Bcc.Add(New Net.Mail.MailAddress(bcc, ""))
                    End If
                Next

                For Each file As String In attachments.Split(";")
                    file = file.Trim
                    If file.Length = 0 Then Continue For
                    If My.Computer.FileSystem.FileExists(file) Then
                        mail.Attachments.Add(New System.Net.Mail.Attachment(file))
                    End If
                Next

                Dim rowTATMAIL1 As DataRow = ABSolution.ASCDATA1.GetDataRow("SELECT * FROM TATMAIL1 WHERE EMAIL_KEY = :PARM1", "V", "SO")

                mail.Subject = emailSubjectText
                If rowTATMAIL1 IsNot Nothing Then
                    EMAIL_LOGO = (rowTATMAIL1.Item("EMAIL_LOGO") & String.Empty).ToString.Trim
                End If

                Dim plainView As Net.Mail.AlternateView = Net.Mail.AlternateView.CreateAlternateViewFromString(documentText)
                Dim htmlView As Net.Mail.AlternateView
                If EMAIL_LOGO <> "" AndAlso ABSolution.ASCMAIN1.Folders.ContainsKey("Images") Then
                    htmlView = Net.Mail.AlternateView.CreateAlternateViewFromString("<img src=cid:logo>" & "<p>" & Replace(documentText & vbCrLf & vbCrLf & SEND_FROM_SIGNATURE, vbCrLf, "<br>") & "</p>", Nothing, "text/html")
                    Dim logo As New Net.Mail.LinkedResource(ABSolution.ASCMAIN1.Folders("Images") & "ABS\" & EMAIL_LOGO)
                    logo.ContentId = "logo"
                    htmlView.LinkedResources.Add(logo)
                Else
                    htmlView = Net.Mail.AlternateView.CreateAlternateViewFromString("<p>" & documentText & vbCrLf & vbCrLf & SEND_FROM_SIGNATURE & "</p>", Nothing, "text/html")
                End If

                mail.AlternateViews.Add(plainView)
                mail.AlternateViews.Add(htmlView)

                Dim smtp As New Net.Mail.SmtpClient(ABSolution.ASCMAIN1.rowASTPARM1.Item("AS_PARM_EMAIL_SMTP_IP"), Val(ABSolution.ASCMAIN1.rowASTPARM1.Item("AS_PARM_EMAIL_SMTP_PORT")))
                smtp.Credentials = New System.Net.NetworkCredential(rowTATMAIL1.Item("EMAIL_ACCT_ID"), rowTATMAIL1.Item("EMAIL_ACCT_PWD"))

                smtp.Send(mail)

            Catch ex As Exception
                RecordLogEntry("Send Email: " & ex.Message)
            End Try
        End Sub


#End Region

#Region "DataSet Functions"

        Private Function ClearDataSetTables(ByVal ClearXMTtables As Boolean) As Boolean

            Try

                If testMode Then RecordLogEntry("Enter ClearDataSetTables.")

                If testMode Then RecordLogEntry("Exit ClearDataSetTables.")
                Return True

            Catch ex As Exception
                RecordLogEntry("ClearDataSetTables: " & ex.Message)
                Return False
            End Try

        End Function

        Private Function PrepareDatasetEntries() As Boolean

            Try

                Dim sql As String = String.Empty
                If testMode Then RecordLogEntry("Enter PrepareDatasetEntries.")

                dst = baseClass.clsASCBASE1.dst
                dst.Tables.Clear()

                With dst

                    baseClass.Get_PARM("SOTPARM1")
                    baseClass.Create_Lookup("ASTUSER1")
                    baseClass.Create_Lookup("ARTCUST1")
                    baseClass.Create_Lookup("ARTCUST2")

                End With

                If testMode Then RecordLogEntry("Exit PrepareDatasetEntries.")
                Return True

            Catch ex As Exception
                RecordLogEntry("PrepareDatasetEntries: " & ex.Message)
                Return False
            End Try

        End Function

        Private Sub UpdateDataSetTables(ByVal invoicesNumbers As String, ByVal invType As String)

            Dim sql As String = String.Empty

            With baseClass
                Try
                    If testMode Then RecordLogEntry("Enter UpdateDataSetTables.")

                    If invoicesNumbers.Length = 0 Then Exit Sub

                    .BeginTrans()

                    Select Case invType

                        Case "D", "E"
                            sql = "Update SOTINVH1 " _
                            & " Set INV_DPD_PRINT_IND = 'P', INV_DPD_PRINT_DATE = SYSDATE, INV_DPD_PRINT_OPER = :PARM1" _
                            & " where INV_NO in (" & invoicesNumbers & ")"
                            ABSolution.ASCDATA1.ExecuteSQL(sql, "V", New Object() {ABSolution.ASCMAIN1.USER_ID})

                        Case "C"
                            sql = "Update SOTRTRN1 " _
                            & " Set INV_PRINTED = '1', INV_PRINTED_DATE = SYSDATE, INV_PRINTED_BY = :PARM1" _
                            & " where RTRN_NO in (" & invoicesNumbers & ")"
                            ABSolution.ASCDATA1.ExecuteSQL(sql, "V", New Object() {ABSolution.ASCMAIN1.USER_ID})

                    End Select


                    .CommitTrans()
                    If testMode Then RecordLogEntry("Exit UpdateDataSetTables.")

                Catch ex As Exception
                    .Rollback()
                    RecordLogEntry("UpdateDataSetTables  : " & ex.Message)
                End Try
            End With

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

                If Not My.Computer.FileSystem.DirectoryExists(logdirectory) Then
                    My.Computer.FileSystem.CreateDirectory(logdirectory)
                End If


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


