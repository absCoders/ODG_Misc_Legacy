Imports System.Data

Public Class Receivings

    Dim WH_OPER_ID As String = ""
    Dim curInvPackUpc As String = ""

    Dim scannerService As New scannerService.ScannerService()
    Dim dst As New Data.DataSet
    'Grid will load the PO Information
    'User will scan ICTPOREC.INV_PACK_UPC
    'Show scan qty on left, item code, qty ordered
    'Initial scan qty of 0 increment w/scan
    'Look into changing color of grid rows.
    Dim itemScanMode As Boolean = False
    Dim singleScanMode As Boolean = False

    Private MyAudioController As Symbol.Audio.Controller = Nothing

    'Need service functions to:
    'LoadPO(ICTPOREC.INV_PACK_UPC) returns item info and qtys
    'UpdateReceipt(dst) send scan info to update function to write to POTSCAN tables




    Private Sub Receivings_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim f2 As New Login()
        Dim vals = f2.GetID()
        WH_OPER_ID = vals(0)

        Barcode1.EnableScanner = True

        Dim dgts As New DataGridTableStyle
        dgts.MappingName = "POTORDR2"
        Dim itemCode As New DataGridTextBoxColumn
        Dim itemQty As New DataGridTextBoxColumn
        Dim qtyOnOrder As New DataGridTextBoxColumn
        itemCode.MappingName = "ITEM_CODE"
        itemCode.HeaderText = "Item Code"
        itemCode.Width = 120
        itemQty.MappingName = "SCAN_QTY"
        itemQty.HeaderText = "Qty"
        itemQty.Width = 30
        qtyOnOrder.MappingName = "OPEN_BALANCE"
        qtyOnOrder.HeaderText = "Open Qty"
        qtyOnOrder.Width = 45
        dgts.GridColumnStyles.Add(itemQty)
        dgts.GridColumnStyles.Add(itemCode)
        dgts.GridColumnStyles.Add(qtyOnOrder)
        DataGrid1.TableStyles.Add(dgts)

        Dim SUMMARY As New DataTable("SUMMARY")
        SUMMARY.Columns.Add("SCAN_SUM", System.Type.GetType("System.Int32"))
        SUMMARY.Columns.Add("OPEN_SUM", System.Type.GetType("System.Int32"))
        dst.Tables.Add(SUMMARY)

        Dim dgts2 As New DataGridTableStyle
        dgts2.MappingName = "SUMMARY"
        Dim qtySum As New DataGridTextBoxColumn
        qtySum.MappingName = "OPEN_SUM"
        qtySum.Width = 45

        Dim qtyCount As New DataGridTextBoxColumn
        qtyCount.MappingName = "SCAN_SUM"
        qtyCount.Width = 150
        dgts2.GridColumnStyles.Add(qtyCount)
        dgts2.GridColumnStyles.Add(qtySum)
        'DataGrid2.RowHeadersVisible = False
        DataGrid2.ColumnHeadersVisible = False
        DataGrid2.TableStyles.Add(dgts2)

        DataGrid2.DataSource = dst.Tables("SUMMARY")

        Dim MyDevice As Symbol.Audio.Device = _
        CType(Symbol.StandardForms.SelectDevice.Select( _
        Symbol.Audio.Controller.Title, _
        Symbol.Audio.Device.AvailableDevices), Symbol.Audio.Device)
        MyAudioController = New Symbol.Audio.StandardAudio(MyDevice)

    End Sub

    Private Sub CalculateSum()
        dst.Tables("SUMMARY").Rows.Clear()
        Dim nrow As DataRow = dst.Tables("SUMMARY").NewRow
        nrow.Item("SCAN_SUM") = dst.Tables("POTORDR2").Compute("SUM(SCAN_QTY)", "")
        nrow.Item("OPEN_SUM") = dst.Tables("POTORDR2").Compute("SUM(OPEN_BALANCE)", "")
        dst.Tables("SUMMARY").Rows.Add(nrow)
    End Sub


    Private Sub Barcode1_OnRead(ByVal sender As System.Object, ByVal readerData As Symbol.Barcode.ReaderData) Handles Barcode1.OnRead
        If itemScanMode = False Then
            'txtStatus.Text = "Loading data..."
            LoadPO(readerData.Text)
        Else
            'look for item code
            Dim scanRow As Data.DataRow = UPCRow(readerData.Text)

            If scanRow IsNot Nothing Then
                If singleScanMode Then
                    scanRow.Item("SCAN_QTY") += 1
                Else
                    scanRow.Item("SCAN_QTY") += scanRow.Item("SCAN_MULT")
                End If
            Else
                ErrorBeep()
            End If

        End If
        CalculateSum()
    End Sub


    Private Sub Receivings_Closing(ByVal sender As System.Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles MyBase.Closing
        Barcode1.EnableScanner = False
        MyAudioController.Dispose()
    End Sub


    Private Function UPCRow(ByVal upc As String) As System.Data.DataRow
        If dst.Tables("POTORDR2").Select("ITEM_UPC_CODE='" & upc & "' OR ITEM_OPC_CODE='" & upc & "' OR ITEM_CODE='" & upc & "'").Length > 0 Then
            Return dst.Tables("POTORDR2").Select("ITEM_UPC_CODE='" & upc & "' OR ITEM_OPC_CODE='" & upc & "' OR ITEM_CODE='" & upc & "'")(0)
        End If
        Return Nothing
    End Function

    Private Sub btnLoad_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnLoad.Click
        If itemScanMode Then
            'Call the service to update POTSCAN1 and POTSCAN2 with our data

            'First remove all rows from POTORDR2 that have 0 in SCAN_QTY
            txtStatus.Text = "Updating..."

            For Each row As Data.DataRow In dst.Tables("POTORDR2").Select("SCAN_QTY=0")
                dst.Tables("POTORDR2").Rows.Remove(row)
            Next

            If dst.Tables("POTORDR2").Rows.Count > 0 Then
                Try
                    Dim rtrnMsg As String = scannerService.UpdatePO2(WH_OPER_ID, curInvPackUpc, dst)
                    If rtrnMsg = "1" Then
                        txtStatus.Text = "Scans updated."
                        curInvPackUpc = ""
                    Else
                        txtStatus.Text = rtrnMsg
                    End If
                Catch ex As Exception
                    txtStatus.Text = "Connection error... try again."
                    Exit Sub
                End Try
            Else
                txtStatus.Text = "Scan or enter UPC."
            End If
            itemScanMode = False

            lblUPC.Text = "UPC"
            txtPONumber.Text = ""

            btnLoad.Text = "Load"
            dst.Reset()
            Dim SUMMARY As New DataTable("SUMMARY")
            SUMMARY.Columns.Add("SCAN_SUM", System.Type.GetType("System.Int32"))
            SUMMARY.Columns.Add("OPEN_SUM", System.Type.GetType("System.Int32"))
            dst.Tables.Add(SUMMARY)
        Else
            LoadPO(txtPONumber.Text)
        End If
    End Sub

    Private Sub chkScan_CheckStateChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkShowScan.CheckStateChanged
        If chkShowScan.Checked Then
            If itemScanMode Then
                dst.Tables("POTORDR2").DefaultView.RowFilter = "SCAN_QTY > 0"
            End If
        Else
            If itemScanMode Then
                dst.Tables("POTORDR2").DefaultView.RowFilter = ""
            End If
        End If
    End Sub

    Private Sub ErrorBeep()
        'duration in ms, frequency in hz
        Me.MyAudioController.PlayAudio(250, 200)
    End Sub

    Private Function LoadPO(ByVal invPackUPC As String) As String

        Try
            txtStatus.Text = "Loading data..."
            dst.Merge(scannerService.LoadPO(invPackUPC))
            curInvPackUpc = invPackUPC
            'txtStatus.Text = "Data loaded."
            'txtStatus.Text = ex.Message

            'if there's data in the dst we're good to go
            If dst.Tables.Count > 0 Then
                itemScanMode = True

                For Each row As Data.DataRow In dst.Tables("POTORDR2").Select("OPEN_BALANCE=0")
                    dst.Tables("POTORDR2").Rows.Remove(row)
                Next

                dst.Tables("POTORDR2").Columns.Add("SCAN_QTY", System.Type.GetType("System.Int32"))

                For Each row As Data.DataRow In dst.Tables("POTORDR2").Rows
                    row.Item("SCAN_QTY") = 0
                Next

                DataGrid1.DataSource = dst.Tables("POTORDR2")
                itemScanMode = True

                If chkShowScan.Checked Then
                    dst.Tables("POTORDR2").DefaultView.RowFilter = "SCAN_QTY > 0"
                End If
            Else
                MsgBox("No data found")
            End If

            txtStatus.Text = "Scan items now."

            btnLoad.Text = "Save"

            lblUPC.Text = "PO"
            txtPONumber.Text = dst.Tables("POTORDR1").Rows(0).Item("PO_ORDER_NO")

        Catch ex As Exception
            itemScanMode = False
            txtStatus.Text = ex.Message
            Return "0"
        End Try

        Return "1"
    End Function

    Private Sub chkSingleScan_CheckStateChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkSingleScan.CheckStateChanged
        singleScanMode = Not singleScanMode
    End Sub
End Class