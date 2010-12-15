Imports System
Imports System.Windows.Forms
Imports System.Runtime.InteropServices
Imports System.Reflection


Public Class SingleInstance

    <DllImport("coredll.dll", SetLastError:=True)> _
    Public Shared Function CreateMutex(ByVal attr As IntPtr, ByVal own As Boolean, ByVal name As String) As IntPtr
    End Function

    <DllImport("coredll.dll", SetLastError:=True)> _
    Public Shared Function ReleaseMutex(ByVal hMutex As IntPtr) As Boolean
    End Function

    <DllImport("aygshell.dll", SetLastError:=True)> _
    Public Shared Function SHFullScreen(ByVal hWnd As IntPtr, ByVal flags As Integer) As Integer
    End Function

    <DllImport("coredll.dll")> _
    Public Shared Function GetCapture() As IntPtr
    End Function

    <DllImport("coredll.dll")> _
    Public Shared Function FindWindow(ByVal className As String, ByVal windowName As String) As IntPtr
    End Function

    <DllImport("coredll.dll", EntryPoint:="EnableWindow")> Private Shared Function EnableWindow(ByVal hwnd As IntPtr, ByVal bEnable As Boolean) As Boolean
    End Function


    Const ERROR_ALREADY_EXISTS As Int32 = 183
    Private Const SHFS_SHOWTASKBAR As Integer = &H1
    Private Const SHFS_HIDETASKBAR As Integer = &H2

    Public Shared Sub Run(ByVal frm As Form)
        Dim name As String = Assembly.GetExecutingAssembly().GetName().Name
        Dim mutexHandle As IntPtr = CreateMutex(IntPtr.Zero, True, name)
        Dim err As Long = Marshal.GetLastWin32Error()

        If (err <> ERROR_ALREADY_EXISTS) Then
            Application.Run(frm)
        End If
        ReleaseMutex(mutexHandle)
    End Sub


    Private Shared Function SetTaskBarEnabled(ByVal bEnabled As Boolean) As Boolean
        Dim hwnd As IntPtr = FindWindow("HHTaskBar", Nothing)

        If Not hwnd.Equals(IntPtr.Zero) Then
            If bEnabled Then
                Return EnableWindow(hwnd, True)
            Else
                Return EnableWindow(hwnd, False)
            End If
        End If
        Return True
    End Function

    Private Shared Function SetTaskbarVisible(ByVal hwnd As IntPtr, ByVal visible As Boolean) As Boolean
        'Dim hwnd As IntPtr = FindWindow("HHTaskBar", Nothing)

        If Not hwnd.Equals(IntPtr.Zero) Then
            If visible Then
                Return SHFullScreen(hwnd, SHFS_SHOWTASKBAR)
            Else
                Return SHFullScreen(hwnd, SHFS_HIDETASKBAR)
            End If
        End If
    End Function

    Public Shared Sub ShowTaskBar(ByVal hwnd As IntPtr)
        SetTaskBarEnabled(True)
        SetTaskbarVisible(hwnd, True)
    End Sub

    Public Shared Sub HideTaskBar(ByVal hwnd As IntPtr)
        SetTaskbarVisible(hwnd, False)
        SetTaskBarEnabled(False)
    End Sub

End Class

