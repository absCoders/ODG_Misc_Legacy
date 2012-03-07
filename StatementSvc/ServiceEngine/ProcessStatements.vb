Imports StatementsEmail.Extensions

Namespace StatementEmail

    Public Class StatementEmailer

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
        Private rowGLTPARM1 As DataRow = Nothing
        Private currentPeriod As String = String.Empty

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
                Dim milTime As String = svcConfig.StartStatements
                Dim emailDay As String = (svcConfig.StatementDay & String.Empty).ToUpper.Trim
                Dim sLastTimeExecuted As String = (svcConfig.LastTimeExecuted & String.Empty).ToUpper.Trim

                Dim processStatements As Boolean = True

                If emailDay.Length = 0 Then
                    emailDay = "ALL"
                ElseIf emailDay.Length > 3 Then
                    emailDay = emailDay.Substring(0, 3)
                End If

                If emailDay <> "ALL" Then
                    If DateTime.Now.ToString("ddd").ToUpper <> emailDay Then
                        RecordLogEntry("MainProcess: Invalid day to process statements")
                        processStatements = False
                    End If
                End If

                If (milTime = "0000") Then
                    RecordLogEntry("MainProcess: Start time set 0000, indicates do not send statements")
                    processStatements = False
                ElseIf (milTime.Length <> 4) Then
                    RecordLogEntry("MainProcess: Invalid Military time to start sending statements")
                    processStatements = False
                Else
                    If (CInt(milTime.Substring(0, 2)) < 12) Then
                        milTime = milTime.Substring(0, 2) + ":" + milTime.Substring(2, 2) + "AM"
                    Else
                        milTime = CStr(CInt(milTime.Substring(0, 2)) - 12) + ":" + milTime.Substring(2, 2) + "PM"
                    End If
                End If

                If DateTime.Now.Hour < CDate(milTime).Hour _
                    OrElse DateTime.Now.Minute < CDate(milTime).Minute Then
                    RecordLogEntry("MainProcess: To early to start emailing statements")
                    processStatements = False
                End If


                If IsDate(sLastTimeExecuted) AndAlso processStatements Then
                    Select Case DateDiff(DateInterval.Day, CDate(DateTime.Now.ToString("MM/dd/yyyy")), CDate(CDate(sLastTimeExecuted).ToString("MM/dd/yyyy")))
                        Case 0
                            ' Same day
                            RecordLogEntry("Main Process: Statements already sent today.")
                            processStatements = False

                        Case Is > 0
                            ' Future date
                            RecordLogEntry("Main Process: Date issue in Config XML file.")
                            processStatements = False

                        Case Is < 0
                            RecordLogEntry("Main Process: Ok to send statements.")
                            svcConfig.UpdateConfigNode("LastTimeExecuted", DateTime.Now)
                    End Select
                End If

                If processStatements Then
                    System.Threading.Thread.Sleep(2000)
                    If LogIntoDatabase() Then
                        System.Threading.Thread.Sleep(2000)
                        If InitializeSettings() Then
                            System.Threading.Thread.Sleep(2000)
                            If PrepareDatasetEntries() Then
                                System.Threading.Thread.Sleep(2000)
                                EmailStatementsToCustomers()
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

                Dim svcConfig As New ServiceConfig
                Dim DriveLetter As String = svcConfig.DriveLetter.ToString.ToUpper
                Dim DriveLetterIP As String = svcConfig.DriveLetterIP.ToString.ToUpper
                Dim convert As Boolean = DriveLetter.Length > 0 AndAlso DriveLetterIP.Length > 0

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

                ' Use the Archive Folder from ASTPARM1
                Dim rowASTPARM1 As DataRow = ABSolution.ASCDATA1.GetDataRow("SELECT * FROM ASTPARM1 WHERE AS_PARM_KEY = 'Z'")
                If rowASTPARM1 IsNot Nothing Then
                    ABSolution.ASCMAIN1.Folders("Archive") = rowASTPARM1.Item("AS_PARM_ARCHIVE_FOLDER") & String.Empty
                End If


                If Not ABSolution.ASCMAIN1.Folders("Archive").EndsWith("\") Then
                    ABSolution.ASCMAIN1.Folders("Archive") &= "\"
                End If

                ABSolution.ASCMAIN1.Folders("Images") = "S:\ODG\Images\"

                For Each field As String In New String() {"Images", "Archive"}
                    If convert And ABSolution.ASCMAIN1.Folders(field).StartsWith(DriveLetter) Then
                        ABSolution.ASCMAIN1.Folders(field) = ABSolution.ASCMAIN1.Folders(field).Replace(DriveLetter, DriveLetterIP)
                    End If
                Next

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

        Private Function RecordEvent(ByVal CUST_CODE As String, ByVal CUST_SHIP_TO_NO As String, ByVal EventDescription As String, Optional ByVal SpokeWith As String = "") As Boolean

            Try
                Dim CONV_NO As String = ABSolution.ASCMAIN1.Next_Control_No("ARTCUSTT.CONV_NO")

                dst.Tables("ARTCUSTT").Clear()

                Dim rowARTCUSTT As DataRow = dst.Tables("ARTCUSTT").NewRow
                rowARTCUSTT.Item("CONV_NO") = CONV_NO
                rowARTCUSTT.Item("CUST_CODE") = CUST_CODE
                rowARTCUSTT.Item("CUST_SHIP_TO_NO") = CUST_SHIP_TO_NO
                rowARTCUSTT.Item("DATE_CONV") = DateTime.Now.ToString("MM/dd/yyyy")
                rowARTCUSTT.Item("SPOKE_WITH") = SpokeWith
                rowARTCUSTT.Item("CONV_LOG") = EventDescription
                rowARTCUSTT.Item("INIT_OPER") = ABSolution.ASCMAIN1.USER_ID
                rowARTCUSTT.Item("INIT_DATE") = DateTime.Now
                dst.Tables("ARTCUSTT").Rows.Add(rowARTCUSTT)

                With baseClass
                    Try
                        .BeginTrans()
                        .clsASCBASE1.Update_Record_TDA("ARTCUSTT")
                        .CommitTrans()
                    Catch ex As Exception
                        .Rollback()
                        RecordLogEntry("RecordEvent : " & ex.Message)
                    End Try
                End With

            Catch ex As Exception
                ' Nothing at this time 
            Finally
                dst.Tables("ARTCUSTT").Clear()
            End Try
        End Function

        Private Function EmailStatementsToCustomers() As Int16

            Dim numEmails As Int16 = 0
            Dim numFax As Int16 = 0
            Dim emsg As String = String.Empty

            Dim sql As String = String.Empty

            Try
                If testMode Then RecordLogEntry("Enter EmailStatementsToCustomers.")

                Dim svcConfig As New ServiceConfig
                Dim CCemail As String = (svcConfig.CCEmail & String.Empty).ToUpper.Trim

                Dim DriveLetter As String = svcConfig.DriveLetter.ToString.ToUpper
                Dim DriveLetterIP As String = svcConfig.DriveLetterIP.ToString.ToUpper
                Dim convert As Boolean = DriveLetter.Length > 0 AndAlso DriveLetterIP.Length > 0

                currentPeriod = ABSolution.ASCMAIN1.Period_Calc(ABSolution.ASCMAIN1.CYP, -1)

                ' Process email statements by customer
                baseClass.clsASCBASE1.Fill_Records("ARTSTMT1", New Object() {currentPeriod})

                If dst.Tables("ARTSTMT1").Rows.Count = 0 Then
                    Return 0
                End If

                Dim rowARTCUST1 As DataRow = Nothing
                Dim rowARTCUST2 As DataRow = Nothing
                Dim custEmailaddress As String = String.Empty
                Dim attachments As String = String.Empty

                For Each rowARTSTMT1 As DataRow In dst.Tables("ARTSTMT1").Select("", "SREP_CODE, CUST_STMT_SEND")
                    Dim CUST_CODE As String = rowARTSTMT1.Item("CUST_CODE") & String.Empty
                    Dim customerName As String = rowARTSTMT1.Item("CUST_NAME") & String.Empty

                    Dim statementNumber As String = rowARTSTMT1.Item("STMT_NO") & String.Empty
                    Dim statementFilename As String = "S:\OSG\" & currentPeriod & "\PDF\" & statementNumber & ".PDF"
                    Dim emailSubject As String = "Statement for " & Mid(currentPeriod, 5, 2) & "/" & Mid(currentPeriod, 1, 4) & _
                            " (Acct# " & CUST_CODE & " " & customerName & ")"

                    If convert And statementFilename.StartsWith(DriveLetter) Then
                        statementFilename = statementFilename.Replace(DriveLetter, DriveLetterIP)
                    End If

                    If Not My.Computer.FileSystem.FileExists(statementFilename) Then
                        RecordLogEntry("The statement '" & statementFilename & "' for customer (" & CUST_CODE & ")  " & customerName & " cannot be found.")
                        Continue For
                    End If

                    emsg = String.Empty

                    Select Case rowARTSTMT1.Item("CUST_STMT_SEND") & String.Empty

                        Case "E"
                            If (rowARTSTMT1.Item("CUST_EMAIL") & String.Empty).ToString.Length = 0 Then
                                RecordLogEntry("Customer " & CUST_CODE & " does not have an email address")
                                Continue For
                            End If

                            ' do not BCC the sales rep: & rowARTSTMT1.Item("SREP_EMAIL") & String.Empty
                            ' as per Maria on 2/27/2012
                            EmailDocument(rowARTSTMT1.Item("CUST_EMAIL") & String.Empty, rowARTSTMT1.Item("SREP_EMAIL") & String.Empty, emailSubject, statementFilename, CCemail & ";", emsg)
                            If emsg.Length = 0 Then
                                RecordEvent(CUST_CODE, "", "emailed " & emailSubject, rowARTSTMT1.Item("CUST_EMAIL") & String.Empty)
                                numEmails += 1
                            End If

                        Case "F"
                            If (rowARTSTMT1.Item("CUST_FAX") & String.Empty).ToString.Length = 0 Then
                                RecordLogEntry("Customer " & CUST_CODE & " does not have a fax number")
                                Continue For
                            End If

                            FaxDocument(CUST_CODE, _
                                        statementFilename, _
                                        "Attached is the file that you have requested.", _
                                        CUST_CODE, _
                                         ABSolution.ASCMAIN1.USER_ID, _
                                        ABSolution.ASCMAIN1.rowASTPARM1.Item("AS_PARM_INST_NAME") & "", _
                                        customerName, _
                                        ABSolution.ASCMAIN1.rowASTPARM1.Item("AS_PARM_INST_NAME") & " " & emailSubject, _
                                        rowARTSTMT1.Item("CUST_FAX"), _
                                        rowARTSTMT1.Item("CUST_CONTACT") & String.Empty, _
                                        emsg)

                            If emsg.Length = 0 Then
                                numFax += 1
                            End If

                        Case Else
                            RecordLogEntry("Customer " & CUST_CODE & " does not have a valid Statement Send value")
                            Continue For

                    End Select

                    If emsg.Length > 0 Then
                        emsg = "Customer " & CUST_CODE & " experienced the following error when sending monthly statement: " _
                        & Environment.NewLine & emsg

                        EmailDocument(rowARTSTMT1.Item("SREP_EMAIL") & String.Empty, rowARTSTMT1.Item("SREP_EMAIL") & String.Empty, "Error Sending Statements", "", "", "", emsg)
                    End If

                    emsg = String.Empty

                    ABSolution.ASCDATA1.ExecuteSQL("UPDATE ARTSTMT1 SET CUST_STMT_SENT = SYSDATE WHERE OPS_YYYYPP = :PARM1 AND CUST_CODE = :PARM2", _
                                                   "VV", New Object() {currentPeriod, CUST_CODE})
                Next

            Catch ex As Exception
                RecordLogEntry("EmailStatementsToCustomers: " & ex.Message)

            Finally
                RecordLogEntry(numEmails & " Emails sent out")
                RecordLogEntry(numFax & " Faxes sent out")
            End Try

        End Function

        ''' <summary>
        ''' Sends an email
        ''' </summary>
        ''' <remarks></remarks>
        Private Sub EmailDocument(ByVal emailTo As String, ByVal emailFrom As String, ByVal emailSubjectText As String, ByVal attachments As String, ByVal BCCemail As String, ByRef emsg As String, Optional ByVal emailBody As String = "")

            If emailTo.Length = 0 OrElse emailFrom.Length = 0 Then
                emsg = "'email to' or 'email from' is missing when attempting to send Statements"
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
                    emsg = "Email Statement: " & ex.Message
                    RecordLogEntry(emsg)

                End Try
            End Using

        End Sub

        Private Function FaxDocument(ByVal customerNumber As String, ByVal attachment As String, _
                                     ByVal sendBody As String, ByVal sendCode As String, ByVal sendFrom As String, _
                                     ByVal sendFromName As String, ByVal sendName As String, ByVal sendSubject As String, _
                                     ByVal sendTo As String, ByVal sendToName As String, ByRef emsg As String) As Boolean

            Try

                sendTo = sendTo.Trim
                Dim faxnumber As String = String.Empty
                Dim zMsg As String = String.Empty

                For Each ch As Char In sendTo
                    If Char.IsDigit(ch) Then
                        faxnumber &= ch
                    End If
                Next

                Select Case faxnumber.Length
                    Case 7, 10
                        ' Should be a good number
                    Case 11
                        If Not faxnumber.StartsWith("1") Then
                            RecordLogEntry("The provided fax number (" & faxnumber & ") for customer: " & faxnumber & " is 11 characters and does not begin with a '1'.")
                        End If
                    Case Else
                        RecordLogEntry("The provided fax number (" & faxnumber & ") for customer: " & faxnumber & " does not appear to be a valid telephone number.")
                End Select

                sendTo = faxnumber

                Dim fax As New TAC.TACFAXS1

                fax.fax_Username = ABSolution.ASCMAIN1.rowASTPARM1.Item("AS_PARM_EFAX_USERNAME")
                fax.fax_Password = ABSolution.ASCMAIN1.rowASTPARM1.Item("AS_PARM_EFAX_PASSWORD")
                fax.fax_CoverFile = ABSolution.ASCMAIN1.Folders("Archive") & "eFax\Cover.rtf"
                fax.fax_FaxAttachment = attachment

                fax.SEND_BODY = sendBody
                fax.SEND_CODE = sendCode
                fax.SEND_FROM = sendFrom
                fax.SEND_FROM_NAME = sendFromName
                fax.SEND_NAME = sendName
                fax.SEND_SUBJECT = sendSubject
                fax.SEND_TO = sendTo
                fax.SEND_TO_NAME = sendToName
                fax.SendFax()

                Dim sendLog As String = fax.fax_log.ToString
                Dim sendID As String = fax.fax_transportID
                Dim sendNo As String = ABSolution.ASCMAIN1.Next_Control_No("TATSEND1.SEND_NO")

                Dim UpdateInProcess As Boolean = False
                With baseClass
                    Try

                        dst.Tables("ARTCUSTT").Rows.Clear()
                        dst.Tables("TATCONV1").Rows.Clear()

                        Dim rowARTCUSTT As DataRow = dst.Tables("ARTCUSTT").NewRow
                        rowARTCUSTT.Item("CONV_NO") = ABSolution.ASCMAIN1.Next_Control_No("ARTCUSTT.CONV_NO")
                        rowARTCUSTT.Item("CUST_CODE") = customerNumber
                        rowARTCUSTT.Item("DATE_CONV") = Format(DateTime.Now, "MM/dd/yyyy")
                        rowARTCUSTT.Item("SPOKE_WITH") = sendTo
                        Dim CONV_LOG As String = "Faxed " & sendSubject & vbCrLf & sendLog
                        rowARTCUSTT.Item("CONV_LOG") = Mid(CONV_LOG, 1, 1000)
                        rowARTCUSTT.Item("INIT_OPER") = ABSolution.ASCMAIN1.USER_ID
                        rowARTCUSTT.Item("INIT_DATE") = DateTime.Now
                        rowARTCUSTT.Item("SEND_NO") = sendNo
                        dst.Tables("ARTCUSTT").Rows.Add(rowARTCUSTT)

                        Dim rowTATCONV1 As DataRow = dst.Tables("TATCONV1").NewRow
                        rowTATCONV1.Item("CONV_NO") = ABSolution.ASCMAIN1.Next_Control_No("TATCONV1.CONV_NO")
                        rowTATCONV1.Item("CONV_DATE") = Format(DateTime.Now, "MM/dd/yyyy")
                        rowTATCONV1.Item("CONV_SUBJECT") = sendTo
                        rowTATCONV1.Item("CONV_NOTES") = "Faxed " & sendSubject
                        rowTATCONV1.Item("INIT_OPER") = ABSolution.ASCMAIN1.USER_ID
                        rowTATCONV1.Item("INIT_DATE") = DateTime.Now
                        rowTATCONV1.Item("TABLE_NAME") = "ARTCUST1"
                        rowTATCONV1.Item("TABLE_KEY") = customerNumber
                        rowTATCONV1.Item("SEND_NO") = sendNo
                        dst.Tables("TATCONV1").Rows.Add(rowTATCONV1)

                        .BeginTrans()
                        UpdateInProcess = True
                        .clsASCBASE1.Update_Record_TDA("ARTCUSTT")
                        .clsASCBASE1.Update_Record_TDA("TATCONV1")
                        .CommitTrans()
                        UpdateInProcess = False

                    Catch ex As Exception
                        If UpdateInProcess Then .Rollback()
                        emsg = "Fax Statement: " & ex.Message
                        RecordLogEntry(emsg)
                    End Try

                End With

                Dim FILENAME As String = ABSolution.ASCMAIN1.Folders("Archive") & "eFax\Logs\" & sendNo & ".txt"
                If My.Computer.FileSystem.FileExists(FILENAME) Then
                    My.Computer.FileSystem.DeleteFile(FILENAME)
                End If
                Using SW As New System.IO.StreamWriter(FILENAME)
                    SW.Write(sendLog)
                End Using

                Return True

            Catch ex As Exception
                emsg = "Error faxing to customer: " & customerNumber & " " & ex.Message
                RecordLogEntry(emsg)
                Return False
            Finally

            End Try

        End Function

#End Region

#Region "DataSet Functions"

        Private Function ClearDataSetTables(ByVal ClearXMTtables As Boolean) As Boolean

            Try

                If testMode Then RecordLogEntry("Enter ClearDataSetTables.")
                dst.Tables("ARTSTMT1").Clear()
                dst.Tables("ARTCUSTT").Clear()
                dst.Tables("TATCONV1").Clear()

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

                    sql = "Select ARTSTMT1.*, ARTCUST1.CUST_NAME, SOTSREP1.SREP_EMAIL" _
                        & ", ARTCUST1.CUST_CONTACT, ARTCUST1.CUST_EMAIL, ARTCUST1.CUST_PHONE, ARTCUST1.CUST_FAX" _
                        & ", ARTCUST6.CUST_LAST_PMT_REF, ARTCUST6.CUST_LAST_PMT_DATE, ARTCUST6.CUST_LAST_PMT_AMT" _
                        & " from ARTSTMT1, ARTCUST1, ARTCUST6, SOTSREP1 " _
                        & " where ARTSTMT1.OPS_YYYYPP = :PARM1" _
                        & " and ARTCUST1.CUST_CODE = ARTSTMT1.CUST_CODE" _
                        & " and ARTCUST6.CUST_CODE = ARTSTMT1.CUST_CODE" _
                        & " and (NVL(ARTSTMT1.TOTAL_DUE, 0) > 0 OR NVL(ARTSTMT1.TYP_ECP, 0) <> 0) " _
                        & " and ARTSTMT1.CUST_STMT_SEND IN ('F','E') and ARTSTMT1.CUST_STMT_SENT IS NULL" _
                        & " and ARTCUST1.CUST_CLASS_CODE <> 'B2C'" _
                        & " and ARTSTMT1.SREP_CODE = SOTSREP1.SREP_CODE" _
                        & " and SOTSREP1.SREP_CODE IS NOT NULL"

                    baseClass.Create_TDA(.Tables.Add, "ARTSTMT1", sql, 0, False, "V")

                    baseClass.Create_TDA(.Tables.Add, "ARTCUSTT", "*")
                    baseClass.Create_TDA(.Tables.Add, "TATCONV1", "*")
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


