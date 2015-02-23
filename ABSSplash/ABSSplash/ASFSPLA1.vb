Imports System
Imports System.Diagnostics

Public NotInheritable Class ASFSPLA1

    Public time_to_close As Boolean = False


    Private Sub ASFSPLA1_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Dim absPath As String

        Try
            absPath = Configuration.ConfigurationManager.AppSettings("absExeFilePath")
            Version.Text = "Version " & FileVersionInfo.GetVersionInfo(absPath).FileVersion
            Copyright.Text = String.Format("Copyright ©{0}", Now.Year) & vbCrLf & "Applied Business Systems, Inc."
        Catch
            MsgBox("Cannot find necessary application file(ABS.exe)", MsgBoxStyle.OkOnly, "Application Error")
            End
        End Try


        Timer1.Interval = 3000
        Timer1.Enabled = True
        Timer2.Interval = 300
        Timer2.Enabled = True
    End Sub

    Public Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
        Timer1.Enabled = False

        ValidateAndLaunch()
        Me.Close()
        Me.Dispose()
    End Sub

    Private Sub Timer2_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer2.Tick
        If (UltraProgressBar1.Value <> 100) Then
            UltraProgressBar1.PerformStep()
        Else
            UltraProgressBar1.Value = 0
        End If
    End Sub

    Private Sub ValidateAndLaunch()

        Dim sPath As String = Configuration.ConfigurationManager.AppSettings("baseDirectory")
        If Not sPath.EndsWith("\") Then sPath &= "\"
        Dim objFile As String
        ' Execute Standard Batch File
        objFile = String.Format("{0}.BAT", Configuration.ConfigurationManager.AppSettings("clientCode"))
        If Not My.Computer.FileSystem.FileExists(sPath & objFile) Then
            MsgBox("Cannot find necessary application file(" & objFile & ".)", MsgBoxStyle.OkOnly, "Application Error")
            Exit Sub
        End If

        Launch_Application(objFile, sPath, ProcessWindowStyle.Hidden, True, 20000)

    End Sub

    Private Sub Launch_Application(ByVal objfile As String, ByVal spath As String, _
        Optional ByVal ProcessWindowStyle As ProcessWindowStyle = ProcessWindowStyle.Normal, _
        Optional ByVal Wait As Boolean = True, Optional ByVal Wait_Time As Integer = 0)

        Try
            Using objprocess As Process = New Process()

                objprocess.StartInfo.FileName = objfile
                objprocess.StartInfo.WorkingDirectory = spath
                'objprocess.StartInfo.UseShellExecute = True
                objprocess.StartInfo.WindowStyle = ProcessWindowStyle
                objprocess.Start()

                Dim ABSProcess As Process = Nothing
                If Wait = True And Wait_Time = 0 Then
                    objprocess.WaitForExit()
                ElseIf Wait_Time > 0 Then
                    While (objprocess.HasExited = False)
                        objprocess.WaitForExit(100)

                        Application.DoEvents()
                        If ABSProcess Is Nothing And Process.GetProcessesByName("ABS").Length > 0 Then
                            ABSProcess = Process.GetProcessesByName("ABS")(0)
                        End If

                        If ABSProcess IsNot Nothing Then

                            Exit While
                        End If
                    End While
                End If

                Wait_Time = 15000
                While (Wait_Time > 0)
                    Wait_Time -= 100
                    Threading.Thread.Sleep(100)
                    Application.DoEvents()
                    If time_to_close Then
                        Exit While
                    End If
                End While

                objprocess.Close()
            End Using

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.OkOnly, "Application Error")
            Exit Sub
        End Try
        Me.Close()
    End Sub
End Class