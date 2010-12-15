<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Public Class Receivings
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
        Me.txtPONumber = New System.Windows.Forms.TextBox
        Me.lblUPC = New System.Windows.Forms.Label
        Me.DataGrid1 = New System.Windows.Forms.DataGrid
        Me.btnLoad = New System.Windows.Forms.Button
        Me.chkShowScan = New System.Windows.Forms.CheckBox
        Me.txtStatus = New System.Windows.Forms.TextBox
        Me.chkSingleScan = New System.Windows.Forms.CheckBox
        Me.DataGrid2 = New System.Windows.Forms.DataGrid
        Me.SuspendLayout()
        '
        'Barcode1
        '
        '
        'txtPONumber
        '
        Me.txtPONumber.Location = New System.Drawing.Point(39, 2)
        Me.txtPONumber.Name = "txtPONumber"
        Me.txtPONumber.Size = New System.Drawing.Size(120, 23)
        Me.txtPONumber.TabIndex = 0
        '
        'lblUPC
        '
        Me.lblUPC.Location = New System.Drawing.Point(3, 5)
        Me.lblUPC.Name = "lblUPC"
        Me.lblUPC.Size = New System.Drawing.Size(30, 20)
        Me.lblUPC.Text = "UPC"
        '
        'DataGrid1
        '
        Me.DataGrid1.BackgroundColor = System.Drawing.Color.FromArgb(CType(CType(128, Byte), Integer), CType(CType(128, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.DataGrid1.Location = New System.Drawing.Point(3, 73)
        Me.DataGrid1.Name = "DataGrid1"
        Me.DataGrid1.Size = New System.Drawing.Size(232, 174)
        Me.DataGrid1.TabIndex = 2
        '
        'btnLoad
        '
        Me.btnLoad.Location = New System.Drawing.Point(165, 2)
        Me.btnLoad.Name = "btnLoad"
        Me.btnLoad.Size = New System.Drawing.Size(70, 23)
        Me.btnLoad.TabIndex = 3
        Me.btnLoad.Text = "Load"
        '
        'chkShowScan
        '
        Me.chkShowScan.Location = New System.Drawing.Point(3, 52)
        Me.chkShowScan.Name = "chkShowScan"
        Me.chkShowScan.Size = New System.Drawing.Size(102, 20)
        Me.chkShowScan.TabIndex = 7
        Me.chkShowScan.Text = "Show Scans"
        '
        'txtStatus
        '
        Me.txtStatus.Location = New System.Drawing.Point(3, 27)
        Me.txtStatus.Name = "txtStatus"
        Me.txtStatus.ReadOnly = True
        Me.txtStatus.Size = New System.Drawing.Size(232, 23)
        Me.txtStatus.TabIndex = 8
        Me.txtStatus.Text = "Scan or enter UPC."
        '
        'chkSingleScan
        '
        Me.chkSingleScan.Location = New System.Drawing.Point(119, 52)
        Me.chkSingleScan.Name = "chkSingleScan"
        Me.chkSingleScan.Size = New System.Drawing.Size(100, 20)
        Me.chkSingleScan.TabIndex = 10
        Me.chkSingleScan.Text = "Single Scan"
        '
        'DataGrid2
        '
        Me.DataGrid2.BackgroundColor = System.Drawing.Color.FromArgb(CType(CType(128, Byte), Integer), CType(CType(128, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.DataGrid2.Location = New System.Drawing.Point(3, 245)
        Me.DataGrid2.Name = "DataGrid2"
        Me.DataGrid2.Size = New System.Drawing.Size(232, 21)
        Me.DataGrid2.TabIndex = 12
        '
        'Receivings
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(96.0!, 96.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Dpi
        Me.AutoScroll = True
        Me.ClientSize = New System.Drawing.Size(638, 455)
        Me.Controls.Add(Me.DataGrid2)
        Me.Controls.Add(Me.chkSingleScan)
        Me.Controls.Add(Me.txtStatus)
        Me.Controls.Add(Me.chkShowScan)
        Me.Controls.Add(Me.btnLoad)
        Me.Controls.Add(Me.DataGrid1)
        Me.Controls.Add(Me.lblUPC)
        Me.Controls.Add(Me.txtPONumber)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "Receivings"
        Me.Text = "Receivings"
        Me.WindowState = System.Windows.Forms.FormWindowState.Maximized
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents Barcode1 As Barcode.Barcode
    Friend WithEvents txtPONumber As System.Windows.Forms.TextBox
    Friend WithEvents lblUPC As System.Windows.Forms.Label
    Friend WithEvents DataGrid1 As System.Windows.Forms.DataGrid
    Friend WithEvents btnLoad As System.Windows.Forms.Button
    Friend WithEvents chkShowScan As System.Windows.Forms.CheckBox
    Friend WithEvents txtStatus As System.Windows.Forms.TextBox
    Friend WithEvents chkSingleScan As System.Windows.Forms.CheckBox
    Friend WithEvents DataGrid2 As System.Windows.Forms.DataGrid
End Class
