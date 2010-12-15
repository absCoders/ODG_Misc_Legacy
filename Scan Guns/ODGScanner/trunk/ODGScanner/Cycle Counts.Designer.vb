<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Public Class CycleCount
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
        Me.Barcode1 = New Barcode.Barcode
        Me.DataGrid1 = New System.Windows.Forms.DataGrid
        Me.txtBin = New System.Windows.Forms.TextBox
        Me.lblBin = New System.Windows.Forms.Label
        Me.Button1 = New System.Windows.Forms.Button
        Me.txtStatus = New System.Windows.Forms.TextBox
        Me.CheckBox1 = New System.Windows.Forms.CheckBox
        Me.SuspendLayout()
        '
        'Barcode1
        '
        '
        'DataGrid1
        '
        Me.DataGrid1.BackgroundColor = System.Drawing.Color.FromArgb(CType(CType(128, Byte), Integer), CType(CType(128, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.DataGrid1.Location = New System.Drawing.Point(3, 69)
        Me.DataGrid1.Name = "DataGrid1"
        Me.DataGrid1.Size = New System.Drawing.Size(231, 193)
        Me.DataGrid1.TabIndex = 0
        '
        'txtBin
        '
        Me.txtBin.Location = New System.Drawing.Point(39, 4)
        Me.txtBin.Name = "txtBin"
        Me.txtBin.Size = New System.Drawing.Size(117, 23)
        Me.txtBin.TabIndex = 1
        '
        'lblBin
        '
        Me.lblBin.Location = New System.Drawing.Point(6, 5)
        Me.lblBin.Name = "lblBin"
        Me.lblBin.Size = New System.Drawing.Size(28, 20)
        Me.lblBin.Text = "Bin"
        '
        'Button1
        '
        Me.Button1.Location = New System.Drawing.Point(159, 5)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(72, 20)
        Me.Button1.TabIndex = 3
        Me.Button1.Text = "Start Bin"
        '
        'txtStatus
        '
        Me.txtStatus.Location = New System.Drawing.Point(3, 38)
        Me.txtStatus.Name = "txtStatus"
        Me.txtStatus.ReadOnly = True
        Me.txtStatus.Size = New System.Drawing.Size(228, 23)
        Me.txtStatus.TabIndex = 6
        '
        'CheckBox1
        '
        Me.CheckBox1.Checked = True
        Me.CheckBox1.CheckState = System.Windows.Forms.CheckState.Checked
        Me.CheckBox1.Location = New System.Drawing.Point(156, 38)
        Me.CheckBox1.Name = "CheckBox1"
        Me.CheckBox1.Size = New System.Drawing.Size(82, 20)
        Me.CheckBox1.TabIndex = 8
        Me.CheckBox1.Text = "Continue"
        '
        'CycleCount
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(96.0!, 96.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Dpi
        Me.AutoScroll = True
        Me.ClientSize = New System.Drawing.Size(638, 455)
        Me.Controls.Add(Me.CheckBox1)
        Me.Controls.Add(Me.txtStatus)
        Me.Controls.Add(Me.Button1)
        Me.Controls.Add(Me.lblBin)
        Me.Controls.Add(Me.txtBin)
        Me.Controls.Add(Me.DataGrid1)
        Me.KeyPreview = True
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "CycleCount"
        Me.Text = "Cycle Counts"
        Me.WindowState = System.Windows.Forms.FormWindowState.Maximized
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents Barcode1 As Barcode.Barcode
    Friend WithEvents DataGrid1 As System.Windows.Forms.DataGrid
    Friend WithEvents txtBin As System.Windows.Forms.TextBox
    Friend WithEvents lblBin As System.Windows.Forms.Label
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents txtStatus As System.Windows.Forms.TextBox
    Friend WithEvents CheckBox1 As System.Windows.Forms.CheckBox

End Class
