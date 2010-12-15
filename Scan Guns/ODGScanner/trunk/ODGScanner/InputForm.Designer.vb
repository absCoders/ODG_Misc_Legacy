<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Public Class InputForm
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing AndAlso components IsNot Nothing Then
            components.Dispose()
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer
    Private mainMenu1 As System.Windows.Forms.MainMenu

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.mainMenu1 = New System.Windows.Forms.MainMenu
        Me.txtInput = New System.Windows.Forms.TextBox
        Me.lblText = New System.Windows.Forms.Label
        Me.btnOK = New System.Windows.Forms.Button
        Me.SuspendLayout()
        '
        'txtInput
        '
        Me.txtInput.Location = New System.Drawing.Point(3, 32)
        Me.txtInput.Name = "txtInput"
        Me.txtInput.PasswordChar = Global.Microsoft.VisualBasic.ChrW(42)
        Me.txtInput.Size = New System.Drawing.Size(132, 23)
        Me.txtInput.TabIndex = 0
        '
        'lblText
        '
        Me.lblText.Location = New System.Drawing.Point(3, 9)
        Me.lblText.Name = "lblText"
        Me.lblText.Size = New System.Drawing.Size(100, 20)
        Me.lblText.Text = "Enter input:"
        '
        'btnOK
        '
        Me.btnOK.Location = New System.Drawing.Point(141, 32)
        Me.btnOK.Name = "btnOK"
        Me.btnOK.Size = New System.Drawing.Size(29, 23)
        Me.btnOK.TabIndex = 2
        Me.btnOK.Text = "OK"
        '
        'InputForm
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(96.0!, 96.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Dpi
        Me.AutoScroll = True
        Me.ClientSize = New System.Drawing.Size(173, 67)
        Me.Controls.Add(Me.btnOK)
        Me.Controls.Add(Me.lblText)
        Me.Controls.Add(Me.txtInput)
        Me.Location = New System.Drawing.Point(50, 50)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "InputForm"
        Me.Text = "Input"
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents txtInput As System.Windows.Forms.TextBox
    Friend WithEvents lblText As System.Windows.Forms.Label
    Friend WithEvents btnOK As System.Windows.Forms.Button
End Class
