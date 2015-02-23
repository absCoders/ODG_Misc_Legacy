Imports System
Imports System.Diagnostics
Imports System.Security
Imports System.Security.Permissions

Public Class ASFSPLAS

    Dim Testing As Boolean = False

    'TODO: This form can easily be set as the splash screen for the application by going to the "Application" tab
    '  of the Project Designer ("Properties" under the "Project" menu).
    Private Sub ASFSPLA2_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        'Set up the dialog text at runtime according to the application's assembly information.  

        'TODO: Customize the application's assembly information in the "Application" pane of the project 
        '  properties dialog (under the "Project" menu).

        'Application title
        'If My.Application.Info.Title <> "" Then
        '    ApplicationTitle.Text = My.Application.Info.Title
        'Else
        '    'If the application title is missing, use the application name, without the extension
        '    ApplicationTitle.Text = System.IO.Path.GetFileNameWithoutExtension(My.Application.Info.AssemblyName)
        'End If

        If Testing Then MsgBox(" ApplicationTitle.Text = " & ApplicationTitle.Text)

        'Format the version information using the text set into the Version control at design time as the
        '  formatting string.  This allows for effective localization if desired.
        '  Build and revision information could be included by using the following code and changing the 
        '  Version control's designtime text to "Version {0}.{1:00}.{2}.{3}" or something similar.  See
        '  String.Format() in Help for more information.
        '
        'Version.Text = System.String.Format(Version.Text, My.Application.Info.Version.Major, My.Application.Info.Version.Minor, My.Application.Info.Version.Build, My.Application.Info.Version.Revision)

        'Version.Text = System.String.Format(Version.Text, My.Application.Info.Version.Major, My.Application.Info.Version.Minor)

        lblVersion.Text = "Version: " & Val(My.Application.Info.Version.Major) & "." & Val(My.Application.Info.Version.Minor) & "." & Val(My.Application.Info.Version.Build) & "." & Val(My.Application.Info.Version.Revision)

        'Copyright info
        lblCopyright.Text = My.Application.Info.Copyright

        Timer1.Interval = 3000
        Timer1.Enabled = True
    End Sub

    Public Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
        Timer1.Enabled = False

        If Testing Then MsgBox("In Timer1_Tick", MsgBoxStyle.OkOnly, "Timer1-Tick")

        Call ValidateAndLaunch()
        Me.Close()
        Me.Dispose()
    End Sub

    Private Sub ValidateAndLaunch()

        If Testing Then MsgBox("In CopyAndLaunch")

        Dim sPath As String = My.Application.Info.DirectoryPath & "\"
        Dim objFile As String

        ' Execute Standard Batch File
        objFile = "JHI.BAT"
        If Not My.Computer.FileSystem.FileExists(sPath & "JHI.BAT") Then
            MsgBox("Cannot find necessary application file(" & objFile & ".)", MsgBoxStyle.OkOnly, "Application Error")
            Exit Sub
        End If

        Call Launch_Application(objFile, sPath, ProcessWindowStyle.Hidden, True, 10000)

    End Sub

    Private Sub Launch_Application(ByVal objfile As String, ByVal spath As String, _
        Optional ByVal ProcessWindowStyle As System.Diagnostics.ProcessWindowStyle = ProcessWindowStyle.Normal, _
        Optional ByVal Wait As Boolean = True, Optional ByVal Wait_Time As Integer = 0)

        If Testing Then MsgBox("In Launch_Application")

        Try
            If Testing Then MsgBox("Dim objprocess As Process = New Process()")
            Dim objprocess As System.Diagnostics.Process = New System.Diagnostics.Process()

            If Testing Then MsgBox("objprocess.StartInfo.FileName = objFile: " & objfile)
            objprocess.StartInfo.FileName = objfile

            If Testing Then MsgBox("objprocess.StartInfo.WorkingDirectory = sPath: " & spath)
            objprocess.StartInfo.WorkingDirectory = spath
            'objprocess.StartInfo.UseShellExecute = True

            If Testing Then MsgBox("objprocess.StartInfo.WindowStyle = ProcessWindowStyle.Hidden")
            objprocess.StartInfo.WindowStyle = ProcessWindowStyle

            If Testing Then MsgBox("objprocess.Start()")
            objprocess.Start()

            If Testing Then MsgBox("objprocess.WaitForExit()")
            If Wait = True And Wait_Time = 0 Then
                objprocess.WaitForExit()
            ElseIf Wait = True And Wait_Time > 0 Then
                objprocess.WaitForExit(Wait_Time)
            End If

            If Testing Then MsgBox("objprocess.Close()")
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

    End Sub

End Class
