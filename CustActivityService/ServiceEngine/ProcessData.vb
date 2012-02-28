Imports CustomerActivity.Extensions

Namespace CustomerActivity

    Public Class CustomerActivityEmailer

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

        Private rowTATMAIL1 As DataRow = Nothing
        Private rowASTUSER1_EMAIL_FROM As DataRow = Nothing

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

                ' Place a blank line in file to better see where each call starts.
                RecordLogEntry(String.Empty)
                RecordLogEntry("Enter MainProcess.")

                ' See if it is time to process teh emails
                Dim svcConfig As New ServiceConfig
                Dim milTime As String = svcConfig.EmailStartTime
                Dim emailDay As String = (svcConfig.EmailDay & String.Empty).ToUpper.Trim
                Dim sLastTimeExecuted As String = (svcConfig.LastTimeExecuted & String.Empty).ToUpper.Trim

                Dim emailStats As Boolean = True

                If emailDay.Length = 0 Then
                    emailDay = "ALL"
                ElseIf emailDay.Length > 3 Then
                    emailDay = emailDay.Substring(0, 3)
                End If

                If emailDay <> "ALL" Then
                    If DateTime.Now.ToString("ddd").ToUpper <> emailDay Then
                        RecordLogEntry("MainProcess: Invalid day to process statistics")
                        emailStats = False
                    End If
                End If

                If (milTime = "0000") Then
                    RecordLogEntry("MainProcess: Start time set 0000, indicates do not send statistics")
                    emailStats = False
                ElseIf (milTime.Length <> 4) Then
                    RecordLogEntry("MainProcess: Invalid Military time to start sending statistics")
                    emailStats = False
                Else
                    If (CInt(milTime.Substring(0, 2)) < 12) Then
                        milTime = milTime.Substring(0, 2) + ":" + milTime.Substring(2, 2) + "AM"
                    Else
                        milTime = CStr(CInt(milTime.Substring(0, 2)) - 12) + ":" + milTime.Substring(2, 2) + "PM"
                    End If
                End If

                If DateTime.Now.Hour < CDate(milTime).Hour _
                    OrElse DateTime.Now.Minute < CDate(milTime).Minute Then
                    RecordLogEntry("MainProcess: To early to start emailing statistics")
                    emailStats = False
                End If


                If IsDate(sLastTimeExecuted) AndAlso emailStats Then
                    Select Case DateDiff(DateInterval.Day, CDate(DateTime.Now.ToString("MM/dd/yyyy")), CDate(CDate(sLastTimeExecuted).ToString("MM/dd/yyyy")))
                        Case 0
                            ' Same day
                            RecordLogEntry("Main Process: statistics already sent today.")
                            emailStats = False

                        Case Is > 0
                            ' Future date
                            RecordLogEntry("Main Process: Date issue in Config XML file.")
                            emailStats = False

                        Case Is < 0
                            RecordLogEntry("Main Process: Ok to send statistics.")
                            svcConfig.UpdateConfigNode("LastTimeExecuted", DateTime.Now)
                    End Select
                End If

                If emailStats Then
                    System.Threading.Thread.Sleep(2000)
                    If LogIntoDatabase() Then
                        System.Threading.Thread.Sleep(2000)
                        If InitializeSettings() Then
                            System.Threading.Thread.Sleep(2000)
                            If PrepareDatasetEntries() Then
                                System.Threading.Thread.Sleep(2000)
                                EmailStatisticsToSalesReps()
                            End If
                        End If
                    End If
                End If

                If testMode Then RecordLogEntry("Exit MainProcess.")

            Catch ex As Exception
                RecordLogEntry("MainProcess: " & ex.Message)

            Finally
                RecordLogEntry("Closing Log file.")
                CloseLog()
                emailInProcess = False
            End Try

        End Sub

        Public Sub LogIn()

            ' Start Service every 1 hours.
            ' This logic should have the service start on every hour. I added an extra 2 minutes
            Dim startInMinutes As Integer = ((60 - DateTime.Now.Minute) + 2) * 60000
            Dim hour As Integer = 60 * 60000

            If My.Application.Info.DirectoryPath.ToUpper.StartsWith("C:\VS") Then
                importTimer = New System.Threading.Timer _
                (New System.Threading.TimerCallback(AddressOf MainProcess), Nothing, 3000, hour * 3)
            Else
                importTimer = New System.Threading.Timer _
                    (New System.Threading.TimerCallback(AddressOf MainProcess), Nothing, startInMinutes, hour)
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

                rowTATMAIL1 = ABSolution.ASCDATA1.GetDataRow("SELECT * FROM TATMAIL1 WHERE EMAIL_KEY = 'SO'")

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

                If UCase(My.Application.Info.DirectoryPath) Like "C:\VS\*" Then
                    ABSolution.ASCMAIN1.Running_in_VS = True
                    folder_prefix = "\..\..\..\..\"
                Else
                    ABSolution.ASCMAIN1.Running_in_VS = False
                    folder_prefix = "\..\"
                End If

                ' Force
                ABSolution.ASCMAIN1.CLIENT_CODE = "ODG"

                If Not ABSolution.ASCMAIN1.Folders.ContainsKey("Images") Then
                    ABSolution.ASCMAIN1.Folders.Add("Images", ABSolution.ASCMAIN1.GetPath(My.Application.Info.DirectoryPath & folder_prefix & "Images\"))
                End If

                If Not ABSolution.ASCMAIN1.Folders.ContainsKey("Reports") Then
                    ABSolution.ASCMAIN1.Folders.Add("Reports", ABSolution.ASCMAIN1.GetPath(My.Application.Info.DirectoryPath & folder_prefix & "Reports\"))
                End If

                If Not ABSolution.ASCMAIN1.Folders.ContainsKey("DataSets") Then
                    ABSolution.ASCMAIN1.Folders.Add("DataSets", ABSolution.ASCMAIN1.GetPath(My.Application.Info.DirectoryPath & folder_prefix & "DataSets\"))
                End If

                If Not ABSolution.ASCMAIN1.Folders.ContainsKey("Temp") Then
                    ABSolution.ASCMAIN1.Folders.Add("Temp", ABSolution.ASCMAIN1.GetPath(My.Application.Info.DirectoryPath & folder_prefix & "Temp\"))
                End If

                If Not ABSolution.ASCMAIN1.Folders.ContainsKey("Work") Then
                    ABSolution.ASCMAIN1.Folders.Add("Work", ABSolution.ASCMAIN1.GetPath(My.Application.Info.DirectoryPath & folder_prefix & "Work\"))
                End If

                If Not ABSolution.ASCMAIN1.Folders.ContainsKey("bin") Then
                    ABSolution.ASCMAIN1.Folders.Add("bin", ABSolution.ASCMAIN1.GetPath(My.Application.Info.DirectoryPath & folder_prefix & "bin\"))
                End If

                If Not ABSolution.ASCMAIN1.Folders.ContainsKey("Help") Then
                    ABSolution.ASCMAIN1.Folders.Add("Help", ABSolution.ASCMAIN1.GetPath(My.Application.Info.DirectoryPath & folder_prefix & "Help\"))
                End If

                If Not ABSolution.ASCMAIN1.Folders.ContainsKey("Archive") Then
                    ABSolution.ASCMAIN1.Folders.Add("Archive", ABSolution.ASCMAIN1.GetPath(My.Application.Info.DirectoryPath & folder_prefix & "Archive\"))
                End If

                If Not ABSolution.ASCMAIN1.Folders.ContainsKey("Attach") Then
                    ABSolution.ASCMAIN1.Folders.Add("Attach", ABSolution.ASCMAIN1.GetPath(My.Application.Info.DirectoryPath & folder_prefix & "Attach\"))
                End If

                If Not ABSolution.ASCMAIN1.Folders.ContainsKey("root") Then
                    ABSolution.ASCMAIN1.Folders.Add("root", ABSolution.ASCMAIN1.GetPath(My.Application.Info.DirectoryPath & folder_prefix))
                End If

                If Not ABSolution.ASCMAIN1.Folders.ContainsKey("Oracle") Then
                    ABSolution.ASCMAIN1.Folders.Add("Oracle", "C:\oracle\product\10.2.0\Client_1\")
                End If

                ABSolution.ASCMAIN1.ActiveForm = baseClass

                ABSolution.ASCMAIN1.Folders("Images") = "S:\ODG\Images\"

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

        Private Function EmailStatisticsToSalesReps() As Int16

            Dim numEmails As Int16 = 0

            Try
                If testMode Then RecordLogEntry("Enter EmailStatisticsToSalesReps.")

                Dim svcConfig As New ServiceConfig
                Dim CCemail As String = (svcConfig.CCEmail & String.Empty).ToUpper.Trim

                ' Process email statements by customer
                baseClass.clsASCBASE1.Fill_Records("ARTCUST1", New Object() {DateAdd(DateInterval.Day, -7, DateTime.Now)})

                If dst.Tables("ARTCUST1").Rows.Count = 0 Then
                    Return 0
                End If

                Dim customerCode As String = String.Empty
                Dim customerName As String = String.Empty
                Dim emailSubject As String = "Customer(s) with no activity since: " & DateAdd(DateInterval.Day, -7, DateTime.Now).ToString("MM/dd/yyyy")
                Dim emailBody As String = String.Empty
                Dim salesRepEmail As String = String.Empty
                Dim salesRepCode As String = String.Empty


                For Each rowARTCUST1 As DataRow In dst.Tables("ARTCUST1").Select("", "SREP_CODE, CUST_CODE")

                    If salesRepCode <> rowARTCUST1.Item("SREP_CODE") & String.Empty Then

                        If emailBody.Length > 0 Then
                            emailBody &= CreateHtmlFooter()

                            EmailDocument(salesRepEmail, salesRepEmail, emailSubject, String.Empty, CCemail, emailBody)
                            numEmails += 1
                            RecordLogEntry("Email sent to sales rep: " & salesRepCode)
                        End If

                        emailBody = CreateHtmlHeader()

                        salesRepCode = rowARTCUST1.Item("SREP_CODE") & String.Empty
                        salesRepEmail = rowARTCUST1.Item("SREP_EMAIL") & String.Empty
                    End If

                    customerCode = rowARTCUST1.Item("CUST_CODE") & String.Empty
                    customerName = rowARTCUST1.Item("CUST_NAME") & String.Empty

                    emailBody &= CreateHtmlDetail(customerCode, customerName, _
                                                  rowARTCUST1.Item("CUST_LAST_ORDR_NO"), _
                                                  CDate(rowARTCUST1.Item("CUST_LAST_ORDR_DATE")).ToString("MM/dd/yyyy"), _
                                                  Format(Val(rowARTCUST1.Item("CUST_LAST_ORDR_AMT") & String.Empty), "#,##0.00"), _
                                                  Format(Val(rowARTCUST1.Item("CUST_SALES_MTD") & String.Empty), "#,##0.00"), _
                                                  Format(Val(rowARTCUST1.Item("CUST_SALES_YTD") & String.Empty), "#,##0.00"))

                Next

                If emailBody.Length > 0 Then
                    emailBody &= CreateHtmlFooter()

                    EmailDocument(salesRepEmail, salesRepEmail, emailSubject, String.Empty, CCemail, emailBody)
                    numEmails += 1
                End If

            Catch ex As Exception
                RecordLogEntry("EmailStatisticsToSalesReps: " & ex.Message)

            Finally
                RecordLogEntry(numEmails & " Emails sent out")

            End Try

        End Function

        Private Function CreateHtmlHeader() As String
            Dim htmlHeader As String = String.Empty

            htmlHeader = "<html>" & Environment.NewLine
            htmlHeader &= "<body>" & Environment.NewLine
            htmlHeader &= "<style> th { text-align:left; } </style>"
            htmlHeader &= "<style> td { text-align:left; } </style>"

            htmlHeader &= "<table cellpadding=""5"" cellspacing=""0"" border=""0"">" & Environment.NewLine
            htmlHeader &= "<tr>" & Environment.NewLine
            htmlHeader &= "<th>Customer</th>" & Environment.NewLine
            htmlHeader &= "<th>Name</th>" & Environment.NewLine
            htmlHeader &= "<th>Order No</th>" & Environment.NewLine
            htmlHeader &= "<th>Order Date</th>" & Environment.NewLine
            htmlHeader &= "<th>Order Total</th>" & Environment.NewLine
            htmlHeader &= "<th>MTD Sales</th>" & Environment.NewLine
            htmlHeader &= "<th>YTD Sales</th>" & Environment.NewLine
            htmlHeader &= "</tr>" & Environment.NewLine

            Return htmlHeader

        End Function

        Private Function CreateHtmlDetail(ByVal customerCode As String, ByVal customerName As String, _
                                          ByVal orderNumber As String, ByVal orderdate As String, ByVal orderTotal As String, _
                                          ByVal mtdSales As String, ByVal ytdSales As String) As String
            Dim htmlDetail As String

            htmlDetail = "<tr>" & Environment.NewLine
            htmlDetail &= "<td>" & customerCode & "</td>" & Environment.NewLine
            htmlDetail &= "<td>" & customerName & "</td>" & Environment.NewLine
            htmlDetail &= "<td>" & orderNumber & "</td>" & Environment.NewLine
            htmlDetail &= "<td>" & orderdate & "</td>" & Environment.NewLine
            htmlDetail &= "<td style=""text-align:right"">" & orderTotal & "</td>" & Environment.NewLine
            htmlDetail &= "<td style=""text-align:right"">" & mtdSales & "</td>" & Environment.NewLine
            htmlDetail &= "<td style=""text-align:right"">" & ytdSales & "</td>" & Environment.NewLine
            htmlDetail &= "</tr>" & Environment.NewLine

            Return htmlDetail

        End Function


        Private Function CreateHtmlFooter() As String

            Dim htmlFooter As String = String.Empty

            htmlFooter = "</table>" & Environment.NewLine
            htmlFooter &= "</body>" & Environment.NewLine
            htmlFooter &= "</html>" & Environment.NewLine

            Return htmlFooter

        End Function

        ''' <summary>
        ''' Sends an email
        ''' </summary>
        ''' <remarks></remarks>
        Private Sub EmailDocument(ByVal emailTo As String, ByVal emailFrom As String, ByVal emailSubjectText As String, ByVal attachments As String, ByVal BCCemail As String, ByVal emailBody As String)

            If emailTo.Length = 0 OrElse emailFrom.Length = 0 Then
                Exit Sub
            End If

            Dim SEND_FROM_SIGNATURE As String = String.Empty
            Dim EMAIL_LOGO As String = String.Empty
            Dim emailCC As String = String.Empty
            Dim emailBCC As String = BCCemail

            If rowASTUSER1_EMAIL_FROM IsNot Nothing Then
                SEND_FROM_SIGNATURE = _
                  rowASTUSER1_EMAIL_FROM.Item("USER_NAME") & vbCrLf _
                & IIf(rowASTUSER1_EMAIL_FROM.Item("USER_TITLE") & "" <> "", rowASTUSER1_EMAIL_FROM.Item("USER_TITLE") & vbCrLf, "") _
                & IIf(rowASTUSER1_EMAIL_FROM.Item("USER_COMPANY") & "" <> "", rowASTUSER1_EMAIL_FROM.Item("USER_COMPANY") & vbCrLf, "") _
                & "Tel: " & ABSolution.ASCMAIN1.FormatTel(rowASTUSER1_EMAIL_FROM.Item("USER_TELEPHONE") & "", rowASTUSER1_EMAIL_FROM.Item("USER_EXT") & "") & vbCrLf _
                & IIf(rowASTUSER1_EMAIL_FROM.Item("USER_FAX") & "" <> "", "Fax: " & ABSolution.ASCMAIN1.FormatTel(rowASTUSER1_EMAIL_FROM.Item("USER_FAX") & "") & vbCrLf, "") _
                & rowASTUSER1_EMAIL_FROM.Item("USER_EMAIL") & vbCrLf
            End If

            Using mail As New Net.Mail.MailMessage()
                Try
                    mail.From = New Net.Mail.MailAddress(emailFrom, "")
                    emailTo = emailTo.Replace(" ", ";")

                    For Each sendTo As String In emailTo.Split(";")
                        If sendTo.Length > 0 Then
                            mail.To.Add(New Net.Mail.MailAddress(sendTo, ""))
                        End If
                    Next

                    emailCC = emailCC.Replace(" ", ";")
                    For Each cc As String In emailCC.Split(";")
                        If cc.Length > 0 Then
                            mail.CC.Add(New Net.Mail.MailAddress(cc, ""))
                        End If
                    Next

                    emailBCC = emailBCC.Replace(" ", ";")
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

                    mail.Subject = emailSubjectText
                    If rowTATMAIL1 IsNot Nothing Then
                        EMAIL_LOGO = (rowTATMAIL1.Item("EMAIL_LOGO") & String.Empty).ToString.Trim
                        If emailBody.Length = 0 Then
                            emailBody = (rowTATMAIL1.Item("EMAIL_BODY") & String.Empty).ToString.Trim
                        End If
                    End If

                    mail.Body = String.Empty

                    Dim plainView As Net.Mail.AlternateView = Net.Mail.AlternateView.CreateAlternateViewFromString(emailBody)
                    Dim htmlView As Net.Mail.AlternateView
                    If EMAIL_LOGO <> "" AndAlso ABSolution.ASCMAIN1.Folders.ContainsKey("Images") Then
                        htmlView = Net.Mail.AlternateView.CreateAlternateViewFromString("<img src=cid:logo>" & "<p>" & Replace(emailBody & vbCrLf & vbCrLf & SEND_FROM_SIGNATURE, vbCrLf, "<br>") & "</p>", Nothing, "text/html")
                        Dim logo As New Net.Mail.LinkedResource(ABSolution.ASCMAIN1.Folders("Images") & "ABS\" & EMAIL_LOGO)
                        logo.ContentId = "logo"
                        htmlView.LinkedResources.Add(logo)
                    Else
                        htmlView = Net.Mail.AlternateView.CreateAlternateViewFromString("<p>" & emailBody & vbCrLf & vbCrLf & SEND_FROM_SIGNATURE & "</p>", Nothing, "text/html")
                    End If

                    mail.AlternateViews.Add(plainView)
                    mail.AlternateViews.Add(htmlView)

                    Dim smtp As New Net.Mail.SmtpClient(ABSolution.ASCMAIN1.rowASTPARM1.Item("AS_PARM_EMAIL_SMTP_IP"), Val(ABSolution.ASCMAIN1.rowASTPARM1.Item("AS_PARM_EMAIL_SMTP_PORT")))
                    smtp.Credentials = New System.Net.NetworkCredential(rowTATMAIL1.Item("EMAIL_ACCT_ID"), rowTATMAIL1.Item("EMAIL_ACCT_PWD"))

                    smtp.Send(mail)

                Catch ex As Exception

                End Try
            End Using

        End Sub

#End Region

#Region "DataSet Functions"

        Private Function ClearDataSetTables(ByVal ClearXMTtables As Boolean) As Boolean

            Try

                If testMode Then RecordLogEntry("Enter ClearDataSetTables.")
                dst.Tables("ARTCUST1").Clear()

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

                    sql = "Select ARTCUST1.CUST_CODE, ARTCUST1.CUST_NAME, ARTCUST1.SREP_CODE, SOTSREP1.SREP_EMAIL" _
                        & ", ARTCUST6.CUST_LAST_ORDR_NO, ARTCUST6.CUST_LAST_ORDR_DATE, ARTCUST6.CUST_LAST_ORDR_AMT" _
                        & ", ARTCUST6.CUST_SALES_MTD, ARTCUST6.CUST_SALES_YTD " _
                        & " from ARTCUST1, ARTCUST6, SOTSREP1 " _
                        & " where ARTCUST6.CUST_CODE = ARTCUST1.CUST_CODE" _
                        & " and ARTCUST1.SREP_CODE = SOTSREP1.SREP_CODE" _
                        & " and ARTCUST1.CUST_STATUS = 'A' " _
                        & " and ARTCUST1.SREP_CODE IS NOT NULL" _
                        & " and SOTSREP1.SREP_EMAIL IS NOT NULL" _
                        & " and SOTSREP1.SREP_STATUS = 'A'" _
                        & " AND NVL(ARTCUST6.CUST_LAST_ORDR_DATE, SYSDATE) < :PARM1"

                    baseClass.Create_TDA(.Tables.Add, "ARTCUST1", sql, 0, False, "D")

                End With

                If testMode Then RecordLogEntry("Exit PrepareDatasetEntries.")
                Return True

            Catch ex As Exception
                RecordLogEntry("PrepareDatasetEntries: " & ex.Message)
                Return False
            End Try

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
            Try
                logStreamWriter.WriteLine(DateTime.Now & ": " & message)
            Catch ex As Exception

            End Try
        End Sub

        Public Sub CloseLog()
            Try
                If logStreamWriter IsNot Nothing Then
                    logStreamWriter.Close()
                    logStreamWriter.Dispose()
                    logStreamWriter = Nothing
                End If
            Catch ex As Exception

            Finally

            End Try
        End Sub

#End Region

    End Class

End Namespace


