Imports System.Configuration.Install
Imports System.ServiceProcess

<System.ComponentModel.RunInstaller(True)> _
Public Class ServiceInstallation
    Inherits Installer

    Public Sub New()
        Dim pi As New ServiceProcessInstaller
        Dim i As New ServiceInstaller

        pi.Account = ServiceAccount.LocalSystem
        i.ServiceName = "EyeconicOrderStatus"
        i.DisplayName = "Eyeconic Order Status Service"
        i.Description = "Sends order status updates to Eyeconic's web service"

        MyBase.Installers.AddRange(New Installer() {pi, i})
    End Sub
End Class