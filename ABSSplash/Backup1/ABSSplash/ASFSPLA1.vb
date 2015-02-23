Imports System
Imports System.Diagnostics
Imports System.Security
Imports System.Security.Permissions
Imports System.Runtime.InteropServices
Imports System.Runtime.Remoting
Imports System.Runtime.Remoting.Channels
Imports System.Runtime.Remoting.Channels.Ipc

Public NotInheritable Class ASFSPLA1

    Public time_to_close As Boolean = False

    Public Interface ICommunicationService
        Sub SaySomething(ByVal text As String)
    End Interface


    Private Sub ASFSPLA1_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Dim absPath As String

        Try
            absPath = My.Application.Info.DirectoryPath & "\bin\ABS.exe" 'path of ABS.exe on network drive - parameterize?
            Version.Text = "Version " & FileVersionInfo.GetVersionInfo(absPath).FileVersion
        Catch
            MsgBox("Cannot find necessary application file(ABS.exe)", MsgBoxStyle.OkOnly, "Application Error")
            End
        End Try

        'Copyright info
        'Copyright.Text = My.Application.Info.Copyright

        Timer1.Interval = 3000
        Timer1.Enabled = True
        Timer2.Interval = 300
        Timer2.Enabled = True
        'Timer2.Enabled = True


    End Sub

    Public Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
        Timer1.Enabled = False

        ValidateAndLaunch()
        Me.Close()
        Me.Dispose()
    End Sub

    Private Sub Timer2_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer2.Tick
        'If Me.UltraPictureBox1.Appearance.AlphaLevel < 255 Then
        '    Me.UltraPictureBox1.Appearance.AlphaLevel = _
        '    CShort(Me.UltraPictureBox1.Appearance.AlphaLevel + 10)
        '    Me.UltraPictureBox1.Appearance.BackColorAlpha = Infragistics.Win.Alpha.UseAlphaLevel
        'End If
        If (UltraProgressBar1.Value <> 100) Then
            UltraProgressBar1.PerformStep()
        Else
            UltraProgressBar1.Value = 0
        End If
    End Sub

    Private Sub ValidateAndLaunch()

        Dim sPath As String = My.Application.Info.DirectoryPath & "\" ' & "\bin\"
        Dim objFile As String
        ' Execute Standard Batch File
        objFile = "SEA.BAT"
        If Not My.Computer.FileSystem.FileExists(sPath & "SEA.BAT") Then
            MsgBox("Cannot find necessary application file(" & objFile & ".)", MsgBoxStyle.OkOnly, "Application Error")
            Exit Sub
        End If

        Launch_Application(objFile, sPath, ProcessWindowStyle.Hidden, True, 20000)

    End Sub

    Private Sub Launch_Application(ByVal objfile As String, ByVal spath As String, _
        Optional ByVal ProcessWindowStyle As System.Diagnostics.ProcessWindowStyle = ProcessWindowStyle.Normal, _
        Optional ByVal Wait As Boolean = True, Optional ByVal Wait_Time As Integer = 0)

        If My.User.Name = "wjz" Then
            Dim ipcCh As IpcChannel
            ipcCh = New IpcChannel("IPChannelName")

            ChannelServices.RegisterChannel(ipcCh, False)
            RemotingConfiguration.RegisterWellKnownServiceType( _
              GetType(CommunicationService), "SreeniRemoteObj", _
                               WellKnownObjectMode.Singleton)
        End If

        Try
            Dim objprocess As System.Diagnostics.Process = New System.Diagnostics.Process()

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
                        'While Me.Focused
                        '    Application.DoEvents()
                        '    System.Threading.Thread.Sleep(50)
                        'End While
                        Exit While
                    End If
                End While
            End If

            Wait_Time = 15000
            While (Wait_Time > 0)
                Wait_Time -= 100
                System.Threading.Thread.Sleep(100)
                Application.DoEvents()
                If time_to_close Then
                    Exit While
                End If
            End While

            objprocess.Close()
            Try
                objprocess.Dispose()
                objprocess.Close()
                objprocess = Nothing
            Catch ex As Exception
                ' Nothing
            End Try

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.OkOnly, "Application Error")
            Exit Sub
        End Try
        Me.Close()
    End Sub

    Public Class CommunicationService
        Inherits MarshalByRefObject
        Implements ASFSPLA1.ICommunicationService

        Public Sub Command(ByVal commandtext As String) Implements ASFSPLA1.ICommunicationService.SaySomething
            If commandtext = "Close" Then
                ASFSPLA1.time_to_close = True
            End If
        End Sub
    End Class
End Class