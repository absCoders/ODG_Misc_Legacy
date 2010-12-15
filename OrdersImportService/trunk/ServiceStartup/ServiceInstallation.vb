Imports System.Configuration.Install
Imports System.ServiceProcess

<System.ComponentModel.RunInstaller(True)> _
Public Class ServiceInstallation
    Inherits Installer

    Public Sub New()
        Dim pi As New ServiceProcessInstaller
        Dim i As New ServiceInstaller

        pi.Account = ServiceAccount.LocalSystem
        i.ServiceName = "Sales Orders Import"
        i.DisplayName = "ABS Services for ABSolution"
        i.Description = "Create Sales Orders based on data in Oracle import tables"

        MyBase.Installers.AddRange(New Installer() {pi, i})
    End Sub
End Class