Public Class orderStatusUpdater
    Implements IDisposable

    Private config As ServiceConfig
    Dim vspService As VSP.ODGOrderStatusServiceClient
    Private oc As oracleClient
    Private logger As Logger

    Public loadedSuccessfully As Boolean


    Public Sub New()
        Try
            config = New ServiceConfig
            logger = New Logger(config.LogDirectory)
            'logger.RecordLogEntry("Logger created")
            oc = New oracleClient(config.ConnectionString)
            'logger.RecordLogEntry("Oracle connected")
            vspService = New VSP.ODGOrderStatusServiceClient
            'logger.RecordLogEntry("Service connection made")
            loadedSuccessfully = True
        Catch ex As Exception
            loadedSuccessfully = False
            If logger IsNot Nothing Then
                logger.RecordLogEntry(ex.Message)
            End If
        End Try
    End Sub

    Public Sub Dispose() Implements IDisposable.Dispose
        oc.Dispose()
        vspService.Close()
        logger.CloseLog()
    End Sub

    Public Sub DoOrderStatusUpdates()

        Dim orderQueue As DataTable = oc.GetDataTable("SELECT ORDR_NO,LAST_DATE FROM " & _
                                                     "(SELECT ORDR_NO,LAST_DATE,MAX(LAST_DATE) OVER (PARTITION BY ORDR_NO) MAXDATE FROM XSTORDRQ WHERE ORDR_SOURCE='Y')" & _
                                                     " WHERE LAST_DATE=MAXDATE")

        If orderQueue IsNot Nothing Then
            For Each queueRow As DataRow In orderQueue.Rows
                Dim vspResponse As VSP.UpdateOrderStatusResponse = UpdateOrderStatus(queueRow.Item("ORDR_NO").ToString())
                If vspResponse.IsSuccess Then
                    oc.ExecuteSQL("DELETE FROM XSTORDRQ WHERE ORDR_NO=:PARM1 AND LAST_DATE <= :PARM2", queueRow.Item("ORDR_NO"), queueRow.Item("LAST_DATE"))
                    logger.RecordLogEntry(queueRow.Item("ORDR_NO").ToString & ":" & vspResponse.Message)
                Else
                    logger.RecordLogEntry(queueRow.Item("ORDR_NO").ToString & ":" & vspResponse.Message & ":" & vspResponse.DetailedMessage)
                End If
            Next
        End If

    End Sub

    Private Function UpdateOrderStatus(ByVal orderNo As String) As VSP.UpdateOrderStatusResponse
        Dim vspResponse As VSP.UpdateOrderStatusResponse = Nothing

        Try
            Dim orderDetails As VSP.CLOrderDetails = CreateCLOrderDetails(orderNo)

            Try
                vspResponse = vspService.UpdateOrderStatus(orderDetails, config.ServicePassPhrase)
            Catch ex As System.ServiceModel.FaultException
                vspResponse = New VSP.UpdateOrderStatusResponse
                vspResponse.IsSuccess = False
                vspResponse.Message = "Web service call failed"
                vspResponse.DetailedMessage = ex.Message & GetDetailIfExisting(ex)
            End Try

        Catch ex As Exception
            vspResponse = New VSP.UpdateOrderStatusResponse
            vspResponse.IsSuccess = False
            vspResponse.Message = "No order status found"
            vspResponse.DetailedMessage = ex.Message

        End Try

        Return vspResponse
    End Function

    Private Function GetDetailIfExisting(ByVal ex As System.ServiceModel.FaultException) As String
        Dim detailString As String = ""
        Dim messageFault As System.ServiceModel.Channels.MessageFault = ex.CreateMessageFault()
        If (messageFault.HasDetail) Then
            detailString = " || " & messageFault.GetDetail(Of String)()
        End If
        Return detailString
    End Function

    Private Function CreateCLOrderDetails(ByVal orderNo As String) As VSP.CLOrderDetails

        Dim orderDetails As New VSP.CLOrderDetails
        Dim orderDataTable As DataTable = oc.GetDataTable( _
                                   "SELECT * FROM XSTORDRS WHERE ORDR_NO=:PARM1", orderNo)

        Dim orderData As DataRow = orderDataTable.Rows(0)

        orderDetails.OrderId = Integer.Parse(orderData.Item("ORDER_ID").ToString)
        orderDetails.OrderStatus = GetStatusType(orderData.Item("ORDR_STATUS").ToString)
        orderDetails.ODGInvoiceNumber = orderData.Item("INV_NO").ToString
        orderDetails.FreightCost = DirectCast(orderData.Item("INV_FREIGHT"), Decimal)
        orderDetails.ShipDate = Now.Date 'If(orderData.Item("SHIP_DATE") Is DBNull.Value, Date.MinValue, orderData.Item("SHIP_DATE"))
        orderDetails.ShippingCarrier = orderData.Item("CARRIER").ToString
        orderDetails.ShippingTrackingNumber = orderData.Item("TRACKING_NO").ToString
        orderDetails.Items = CreateCLOrderItems(orderNo)

        Return orderDetails
    End Function

    Private Function GetStatusType(ByVal orderStatusCode As String) As VSP.CLOrderDetails.CLOrderStatus
        Dim vspStatus As VSP.CLOrderDetails.CLOrderStatus
        Select Case orderStatusCode
            Case "R"
                vspStatus = VSP.CLOrderDetails.CLOrderStatus.InProcess
            Case "P", "O"
                vspStatus = VSP.CLOrderDetails.CLOrderStatus.InProcess
            Case "B"
                vspStatus = VSP.CLOrderDetails.CLOrderStatus.Backordered
            Case "H"
                vspStatus = VSP.CLOrderDetails.CLOrderStatus.Hold
            Case "F"
                vspStatus = VSP.CLOrderDetails.CLOrderStatus.Shipped
            Case "V"
                vspStatus = VSP.CLOrderDetails.CLOrderStatus.Cancelled
        End Select
        Return vspStatus
    End Function

    Private Function CreateCLOrderItems(ByVal orderNo As String) As List(Of VSP.CLOrderDetails.CLOrderItem)

        Dim orderDetails As DataTable = oc.GetDataTable("SELECT * FROM XSTORDRD WHERE ORDR_NO=:PARM1", orderNo)

        Dim itemList As New List(Of VSP.CLOrderDetails.CLOrderItem)

        For Each detailRow As DataRow In orderDetails.Rows
            Dim orderItem As New VSP.CLOrderDetails.CLOrderItem()

            orderItem.ItemId = Convert.ToInt32(detailRow.Item("ITEM_ID"))
            orderItem.ItemCost = Convert.ToDecimal(detailRow.Item("ORDR_UNIT_PRICE"))
            orderItem.ItemStatus = Me.GetLineStatusType(detailRow.Item("ORDR_LINE_STATUS").ToString)
            orderItem.PRD_ID = detailRow.Item("ITEM_CODE").ToString

            itemList.Add(orderItem)
        Next

        Return itemList
    End Function

    Private Function GetLineStatusType(ByVal orderStatusCode As String) As VSP.CLOrderDetails.CLOrderItemStatus
        Dim vspStatus As VSP.CLOrderDetails.CLOrderItemStatus
        Select Case orderStatusCode
            Case "B"
                vspStatus = VSP.CLOrderDetails.CLOrderItemStatus.Backordered
            Case "P", "O"
                vspStatus = VSP.CLOrderDetails.CLOrderItemStatus.InProcess
            Case "H"
                vspStatus = VSP.CLOrderDetails.CLOrderItemStatus.Hold
            Case "F"
                vspStatus = VSP.CLOrderDetails.CLOrderItemStatus.Shipped
            Case "C", "V"
                vspStatus = VSP.CLOrderDetails.CLOrderItemStatus.Cancelled
        End Select
        Return vspStatus
    End Function

End Class
