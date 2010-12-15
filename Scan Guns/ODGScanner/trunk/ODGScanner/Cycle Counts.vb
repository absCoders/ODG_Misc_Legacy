
'RZ** IS NONSTOCK BIN
'IF PROCESSING A NONSTOCK BIN, ITEM_BIN SHOULD BE ''
'ONLY ALLOW ONE PRICE CATEGORY TO BE PROCESSED PER SCAN GROUP
'TAKE PRICE CATEGORY OF FIRST ITEM SCANNED AND ALLOW ONLY THAT



Public Class CycleCount
    Dim scannerService As New scannerService.ScannerService()
    Dim dst As New System.Data.DataSet

    Private MyAudioController As Symbol.Audio.Controller = Nothing

    Dim scanMode As Boolean
    Dim continuousMode As Boolean = False
    Dim forUpdate As String = "0"
    Dim curBin As String
    Dim curPriceCatgy As String = ""

    Dim completedItems As List(Of String)
    Dim latestItem As String

    Dim WH_OPER_ID As String = ""
    Dim WHSE_CODE As String = ""
    Dim keyedQty As String = ""



    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        Dim f2 As New Login()
        Dim vals = f2.GetID()
        WH_OPER_ID = vals(0)
        WHSE_CODE = vals(1)
        'If WHSE_CODE = "003" Then
        CheckBox1.Visible = True
        txtStatus.Width = 140
        'Else
        '    CheckBox1.Visible = False
        'End If
        scanMode = False
        Dim dgts As New DataGridTableStyle
        dgts.MappingName = "Table"
        Dim itemCode As New DataGridTextBoxColumn
        'Dim itemDesc As New DataGridTextBoxColumn
        Dim itemQty As New DataGridTextBoxColumn
        Dim qtyOnHand As New DataGridTextBoxColumn
        itemCode.MappingName = "ITEM_CODE"
        itemCode.HeaderText = "Item Code"
        itemCode.Width = 120
        'itemDesc.MappingName = "ITEM_DESC"
        'itemDesc.HeaderText = "Desc"
        'itemDesc.Width = 120
        itemQty.MappingName = "SCAN_QTY"
        itemQty.HeaderText = "Qty"
        itemQty.Width = 30
        qtyOnHand.MappingName = "WHSE_QTY_ON_HAND"
        qtyOnHand.HeaderText = "OH Qty"
        qtyOnHand.Width = 45
        dgts.GridColumnStyles.Add(itemQty)
        dgts.GridColumnStyles.Add(itemCode)
        'dgts.GridColumnStyles.Add(itemDesc)
        dgts.GridColumnStyles.Add(qtyOnHand)
        DataGrid1.TableStyles.Add(dgts)

        txtStatus.Text = "Enter a bin no."


        Dim MyDevice As Symbol.Audio.Device = _
        CType(Symbol.StandardForms.SelectDevice.Select( _
        Symbol.Audio.Controller.Title, _
        Symbol.Audio.Device.AvailableDevices), Symbol.Audio.Device)
        MyAudioController = New Symbol.Audio.StandardAudio(MyDevice)

        If WHSE_CODE = "003" Then
            Barcode1.EnableScanner = True 'So that DEL can scan a product to load up a BIN
        End If

        txtBin.Focus()
    End Sub

    Private Sub Barcode1_OnRead(ByVal sender As System.Object, ByVal readerData As Symbol.Barcode.ReaderData) Handles Barcode1.OnRead
        Dim itemCode As String = ""

        If readerData.Result = 0 Then '0 is the code for Symbol.Results.SUCCESS
            If curBin = "" And WHSE_CODE = "003" Then
                Try
                    txtBin.Text = scannerService.GetBin(readerData.Text)
                Catch ex As Exception
                    txtStatus.Text = "Connection error - try again."
                    ErrorBeep()
                    Exit Sub
                End Try

                If txtBin.Text <> "" Then
                    Button1_Click_1(Me, EventArgs.Empty)
                End If
                Exit Sub
            End If
            Dim qty As Integer = 1
            Dim curdst As System.Data.DataSet
            Try
                If Val(keyedQty) > 0 Then
                    qty = Val(keyedQty)
                End If
                'don't call service if upc is already in datagrid, just increment locally
                If dst.Tables.Count > 0 Then
                    If dst.Tables(0).Select("ITEM_UPC_CODE='" & readerData.Text & "' or ITEM_OPC_CODE='" & readerData.Text & "' or ITEM_CODE='" & readerData.Text & "'").Length > 0 Then
                        'increment scan_qty for this item and exit sub

                        'For an already scanned item,
                        '(which this is),
                        'ask for confirmation that they want to add to
                        'the qty for this item.
                        UPCRow(readerData.Text).Item("SCAN_QTY") += qty
                        txtStatus.Text = "Scan successful."
                        keyedQty = ""
                        Exit Sub
                    End If
                End If

                'getscandata and merge with dst if it has content

                'At this point we are only looking at new scans
                Try
                    curdst = scannerService.GetItemInfo(readerData.Text, WHSE_CODE)
                Catch ex As Exception
                    txtStatus.Text = "Connection error - try again."
                    ErrorBeep()
                    Exit Sub
                End Try
                'verify that the item belongs to this bin before adding it to our dst
                If curdst IsNot Nothing AndAlso curdst.Tables.Count > 0 AndAlso curdst.Tables(0).Rows.Count > 0 Then
                    If dst.Tables.Count = 0 Then
                        If (curBin.StartsWith("RZ") And curdst.Tables(0).Rows(0).Item("ITEM_BIN") & "" = "") OrElse curdst.Tables(0).Rows(0).Item("ITEM_BIN") IsNot DBNull.Value AndAlso curdst.Tables(0).Rows(0).Item("ITEM_BIN") = curBin Then
                            dst.Merge(curdst)
                            dst.Tables(0).Columns.Add("SCAN_QTY", System.Type.GetType("System.Int32"))
                            DataGrid1.DataSource = dst.Tables(0)
                            dst.Tables(0).Rows(0).Item("SCAN_QTY") = qty
                            txtStatus.Text = "Scan successful."
                            'Set curPriceCatgy for this group of scans.
                            curPriceCatgy = curdst.Tables(0).Rows(0).Item("PRICE_CATGY_CODE")
                        Else
                            'let user know that item does not belong in this bin
                            If curdst.Tables(0).Rows(0).Item("ITEM_BIN") IsNot DBNull.Value Then
                                txtStatus.Text = "Item bin: " & curdst.Tables(0).Rows(0).Item("ITEM_BIN")
                            End If

                            ErrorBeep()
                        End If
                    Else
                        If (curBin.StartsWith("RZ") And curdst.Tables(0).Rows(0).Item("ITEM_BIN") & "" = "") OrElse curdst.Tables(0).Rows(0).Item("ITEM_BIN") IsNot DBNull.Value AndAlso curdst.Tables(0).Rows(0).Item("ITEM_BIN") = curBin Then
                            'dst.Tables(0).Rows.Add(curdst.Tables(0).Rows(0))
                            If curPriceCatgy = "" Then
                                curPriceCatgy = curdst.Tables(0).Rows(0).Item("PRICE_CATGY_CODE")
                            End If
                            If WHSE_CODE = "003" OrElse curdst.Tables(0).Rows(0).Item("PRICE_CATGY_CODE") = curPriceCatgy Then
                                dst.Tables(0).ImportRow(curdst.Tables(0).Rows(0))
                                UPCRow(readerData.Text).Item("SCAN_QTY") = qty
                            Else
                                txtStatus.Text = "Invalid Price Category."
                                ErrorBeep()
                            End If
                        Else
                            'let user know that item does not belong in this bin
                            txtStatus.Text = "Item bin:" & curdst.Tables(0).Rows(0).Item("ITEM_BIN")
                            ErrorBeep()
                        End If
                    End If
                Else
                    ErrorBeep() 'maybe make this a different tone, is an invalid scan
                    txtStatus.Text = "Invalid scan."
                End If


                'Remove any color already on grid
                'Highlight the just scanned row w/color
                If dst.Tables.Count > 0 Then
                    dst.Tables(0).DefaultView.Sort = "ITEM_CODE"
                End If
                'Dim dv As Data.DataView = New Data.DataView(dst.Tables(0))




            Catch ex As Exception
                MsgBox(ex.Message)
            End Try
            keyedQty = ""
        End If
    End Sub

    Private Sub ErrorBeep()
        'duration in ms, frequency in hz


        Try 'prevent crash if Error in audio
            Me.MyAudioController.PlayAudio(250, 200)
        Catch
            txtStatus.Text = "ERROR"
        End Try
    End Sub

    Private Function UPCRow(ByVal upc As String) As System.Data.DataRow
        Return dst.Tables(0).Select("ITEM_UPC_CODE='" & upc & "' OR ITEM_OPC_CODE='" & upc & "' OR ITEM_CODE='" & upc & "'")(0)
    End Function

    Private Sub Form1_Closing(ByVal sender As System.Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles MyBase.Closing
        Barcode1.EnableScanner = False
        MyAudioController.Dispose()
    End Sub


    Private Sub Button1_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        If scanMode Then
            Dim result As String = "0"
            If dst.Tables.Count > 0 Then 'If something has been scanned -
                If dst.Tables(0).Rows.Count > 0 Then
                    Dim bin As String = curBin
                    If curBin.StartsWith("RZ") Then
                        bin = ""
                    End If

                    Try
                        result = scannerService.UpdateItemInfo(dst, bin, curPriceCatgy, WH_OPER_ID, WHSE_CODE, forUpdate)
                    Catch ex As Exception
                        txtStatus.Text = ex.Message
                        ErrorBeep()
                        Exit Sub
                    End Try
                    forUpdate = "0"
                Else
                    EndBin() 'nothing was scanned
                    txtBin.Focus()
                    Exit Sub
                End If

                If result = "1" Then 'update was successful
                    EndBin()
                Else
                    'update was not successful, alert user
                    txtStatus.Text = result
                    'MsgBox(result)
                    ErrorBeep()
                End If
            Else
                EndBin()
            End If
            txtBin.Focus()
        Else
            StartBin()
        End If
    End Sub

    Private Sub StartBin()
        If txtBin.TextLength > 0 Then
            txtBin.Text = txtBin.Text.ToUpper
            Dim result As String = ""
            Try
                result = scannerService.CheckBin(txtBin.Text)
            Catch ex As System.Net.WebException
                txtStatus.Text = "Connection error - try again."
                'MsgBox(ex.Status.ToString())
                'MsgBox(CType(ex.Response, System.Net.HttpWebResponse).StatusCode)
                'MsgBox(ex.Message)
                ErrorBeep()
                Exit Sub
            End Try
            If result = "1" Then
                Barcode1.EnableScanner = True
                scanMode = True
                Button1.Text = "Update"
                curBin = txtBin.Text
                txtBin.ReadOnly = True
                txtStatus.Text = "Ready to scan."
            Else
                Barcode1.EnableScanner = True
                scanMode = True
                Button1.Text = "Update"
                curBin = txtBin.Text
                txtBin.ReadOnly = True
                txtStatus.Text = "Ready to scan."
                dst = scannerService.GetScanData(curBin)
                DataGrid1.DataSource = dst.Tables(0)
                forUpdate = "1"
            End If
        Else
            'let user know they must enter a bin no
        End If
    End Sub

    Private Sub EndBin()
        If dst.Tables.Count > 0 Then
            dst.Tables(0).Clear()
        End If
        curPriceCatgy = ""
        If continuousMode Then
            curBin = NextBin(curBin)
            If curBin = "" Then
                'continuousMode = False
                'CheckBox1.Checked = False
            Else
                txtBin.Text = curBin
                keyedQty = ""
                StartBin()
            End If
        End If

        If Not continuousMode Or curBin = "" Then
            If WHSE_CODE <> "003" Then
                Barcode1.EnableScanner = False
            End If
            scanMode = False
            Button1.Text = "Start Bin"
            txtBin.ReadOnly = False
            curBin = ""
            txtBin.Text = ""
            keyedQty = ""
            txtStatus.Text = "Enter a bin."
        End If
    End Sub

    Private Function NextBin(ByVal bin As String) As String
        If WHSE_CODE = "003" Then
            If bin.EndsWith("7") Then
                Dim nxt = Chr(Asc(bin(bin.Length - 2)) + 1) & "0"
                bin = bin.Substring(0, bin.Length - 2) & nxt
            Else
                Dim nxt = Chr(Asc(bin(bin.Length - 1)) + 1)
                bin = bin.Substring(0, bin.Length - 1) & nxt
            End If
        Else
            Dim nxt = Chr(Asc(bin(bin.Length - 1)) + 1)
            bin = bin.Substring(0, bin.Length - 1) & nxt
            'use service to check ICTPHYX2 for this bin. if not exist blank out bin
            If scannerService.IsBinValid(bin, WHSE_CODE) <> "1" Then
                bin = ""
            End If
        End If
        Return bin
    End Function

    Private Sub Form1_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles MyBase.KeyPress
        If scanMode Then
            'lblStatus.Text = e.KeyChar
            If e.KeyChar >= "0" And e.KeyChar <= "9" Then
                keyedQty &= e.KeyChar
                txtStatus.Text = "Qty to add: " & keyedQty
            End If

            If e.KeyChar = Chr(8) Then 'backspace key erases any qty entered
                If DataGrid1.CurrentRowIndex >= 0 Then
                    Dim cm As CurrencyManager
                    cm = Me.BindingContext(Me.DataGrid1.DataSource)
                    If Not cm Is Nothing Then cm.RemoveAt(DataGrid1.CurrentRowIndex)
                    Dim rmRow As Data.DataRow = Nothing
                    For Each row As Data.DataRow In dst.Tables(0).Rows
                        If row.RowState = Data.DataRowState.Deleted Then
                            rmRow = row
                        End If
                    Next
                    If rmRow IsNot Nothing Then
                        dst.Tables(0).Rows.Remove(rmRow)
                    End If
                End If
                keyedQty = ""
                txtStatus.Text = "Ready to scan."
            End If

            'If e.KeyChar = "*" Then
            '    'Cubby complete.
            '    'If another scan of a 'completed' item comes through,
            '    'Ask for confirmation of the scan.
            '    completedItems.Add(latestItem)
            'End If
        End If

        If e.KeyChar = Chr(13) Then
            'click update/start
            Button1_Click_1(Me, EventArgs.Empty)
        End If
    End Sub

    Private Sub txtBin_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtBin.KeyPress
        If e.KeyChar = Chr(13) Then

        End If
    End Sub

    Private Sub CheckBox1_CheckStateChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox1.CheckStateChanged
        continuousMode = Not continuousMode
    End Sub
End Class