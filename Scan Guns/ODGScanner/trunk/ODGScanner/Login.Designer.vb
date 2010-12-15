<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Public Class Login
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

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.lblOperID = New System.Windows.Forms.Label
        Me.txtOperID = New System.Windows.Forms.TextBox
        Me.lblWarehouse = New System.Windows.Forms.Label
        Me.cmbWarehouse = New System.Windows.Forms.ComboBox
        Me.btnBegin = New System.Windows.Forms.Button
        Me.SuspendLayout()
        '
        'lblOperID
        '
        Me.lblOperID.Location = New System.Drawing.Point(3, 21)
        Me.lblOperID.Name = "lblOperID"
        Me.lblOperID.Size = New System.Drawing.Size(92, 20)
        Me.lblOperID.Text = "Operator ID:"
        '
        'txtOperID
        '
        Me.txtOperID.Location = New System.Drawing.Point(91, 18)
        Me.txtOperID.MaxLength = 4
        Me.txtOperID.Name = "txtOperID"
        Me.txtOperID.Size = New System.Drawing.Size(100, 23)
        Me.txtOperID.TabIndex = 1
        '
        'lblWarehouse
        '
        Me.lblWarehouse.Location = New System.Drawing.Point(3, 56)
        Me.lblWarehouse.Name = "lblWarehouse"
        Me.lblWarehouse.Size = New System.Drawing.Size(79, 20)
        Me.lblWarehouse.Text = "Warehouse:"
        '
        'cmbWarehouse
        '
        Me.cmbWarehouse.Location = New System.Drawing.Point(91, 53)
        Me.cmbWarehouse.Name = "cmbWarehouse"
        Me.cmbWarehouse.Size = New System.Drawing.Size(100, 23)
        Me.cmbWarehouse.TabIndex = 3
        '
        'btnBegin
        '
        Me.btnBegin.Location = New System.Drawing.Point(91, 91)
        Me.btnBegin.Name = "btnBegin"
        Me.btnBegin.Size = New System.Drawing.Size(72, 20)
        Me.btnBegin.TabIndex = 4
        Me.btnBegin.Text = "Begin"
        '
        'Login
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(96.0!, 96.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Dpi
        Me.AutoScroll = True
        Me.ClientSize = New System.Drawing.Size(204, 123)
        Me.Controls.Add(Me.btnBegin)
        Me.Controls.Add(Me.cmbWarehouse)
        Me.Controls.Add(Me.lblWarehouse)
        Me.Controls.Add(Me.txtOperID)
        Me.Controls.Add(Me.lblOperID)
        Me.Location = New System.Drawing.Point(20, 50)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "Login"
        Me.Text = "Enter ID"
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents lblOperID As System.Windows.Forms.Label
    Friend WithEvents txtOperID As System.Windows.Forms.TextBox
    Friend WithEvents lblWarehouse As System.Windows.Forms.Label
    Friend WithEvents cmbWarehouse As System.Windows.Forms.ComboBox
    Friend WithEvents btnBegin As System.Windows.Forms.Button
End Class
