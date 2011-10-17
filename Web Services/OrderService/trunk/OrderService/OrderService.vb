Imports System.Xml.Serialization
' NOTE: If you change the class name "OrderRequest" here, you must also update the reference to "OrderRequest" in App.config.

#Const TEST = True

<ServiceBehavior(Namespace:="http://www.absolution.com/schemas/ODG/Services")> _
Public Class OrderService
    Implements IOrderService

    Private oc As oracleClient
    Private xmlns As String = "http://www.absolution.com/schemas/ODG/Services"

#If TEST Then
    Private logDirectory = "C:\OrderRequests\_Test\"
    Private connectionString = "Data Source=TST;User ID=TST;Password=TST;pooling=true"
    Private servicePassPhrase = "passphrase"
#Else
    Private logDirectory = "C:\OrderRequests\"
    Private connectionString = "Data Source=ODG;User ID=ODG;Password=ODG;pooling=true"
    Private servicePassPhrase = "c0ntActl3nsEs"
#End If

    Private validSourceTypes As List(Of String) = New List(Of String)(New String() {"AnyLens", "Cohens"})
    Private orderSources As Dictionary(Of String, String)

#Region "OrderRequestFunctions"

    Public Function PlaceOrder(ByVal orderRequest As OrderRequest, ByVal passPhrase As String) As OrderResponse Implements IOrderService.PlaceOrder

        orderSources = New Dictionary(Of String, String)
        orderSources.Add("AnyLens", "A")
        orderSources.Add("Cohens", "C")
        Dim orderResponse = PerformRequestValidation(orderRequest, passPhrase)

        If Not orderResponse.errorsOccurred Then
            logDirectory &= orderRequest.Source.type & "\"
            oc = New oracleClient(connectionString)

            For Each order As Order In orderRequest.Source.Orders
                SetOrderDefaults(order)
                Dim orderResult As OrderResult = PerformOrderValidation(order, orderRequest.Source.type)

                SaveXML(order, orderResult)

                If Not orderResult.orderHasErrors Then 'No validation errors, attempt save to DB
                    Try
                        WriteOrderToDB(order, orderRequest.Source.type)
                        ArchiveOrderXML(order.id)
                    Catch ex As Exception
                        LogError(order.id, ex)
                        orderResult.AddOrderError(String.Format("Error saving order to database: {0}", ex.Message))
                        orderResponse.errorsOccurred = True
                    End Try
                Else
                    MoveXMLtoError(order.id, ErrorType.Validation)
                    orderResponse.errorsOccurred = True
                End If

                orderResponse.OrderResults.Add(orderResult)
            Next

            oc.Dispose()
        End If

        Return orderResponse
    End Function

    Private Sub SetOrderDefaults(ByVal order As Order)
        If String.IsNullOrEmpty(order.Shipping.TaxShipping) Then
            order.Shipping.TaxShipping = "N"
        End If

        order.Shipping.Telephone = order.Shipping.Telephone.Left(10)
    End Sub

#End Region

#Region "OrderStatusFunctions"

    Public Function GetOrderStatus(ByVal statusRequest As StatusRequest, ByVal passPhrase As String) As StatusResponse Implements IOrderService.GetOrderStatus
        Dim statusResponse As New StatusResponse()
        If IsPassPhraseValid(passPhrase) Then
            statusResponse.type = "OrderStatus"

            If statusRequest.type = "OrderStatus" Then
                oc = New oracleClient(connectionString)

                Dim orderLineSource As String = oc.GetDataValue("SELECT ORDR_LINE_SOURCE FROM XMTXREF1 WHERE XML_ORDR_SOURCE=:PARM1", New Object() {statusRequest.Source.type})


                For Each item As StatusRequestItem In statusRequest.Source.Items.itemList
                    Dim responseItem As New StatusResponseItem
                    responseItem.id = item.id

                    Dim itemSql As String = "SELECT" _
                                        & " O2.ORDR_LINE_STATUS,O2.ORDR_QTY,O2.ORDR_QTY_SHIP,O2.ORDR_QTY_BACK,O2.ORDR_QTY_CANC,O2.ORDR_QTY_OPEN,O2.ORDR_QTY_PICK,O2.ORDR_UNIT_PRICE,O1.LAST_DATE, I1.INV_NO, I1.INV_DATE, I1.SHIP_REF,I1.SHIP_VIA_CODE,NVL(I1.INV_FREIGHT,0) INV_FREIGHT" _
                                        & " FROM" _
                                        & " SOTORDR1 O1" _
                                        & " JOIN" _
                                        & " SOTORDR2 O2 ON (O1.ORDR_NO=O2.ORDR_NO)" _
                                        & " LEFT JOIN" _
                                        & " SOTINVH1 I1 ON (I1.ORDR_NO=O1.ORDR_NO)" _
                                        & " LEFT JOIN" _
                                        & " SOTINVH2 I2 ON (I1.INV_NO=I2.INV_NO AND O2.ORDR_LNO=I2.INV_LNO)" _
                                        & " WHERE" _
                                        & " O2.CUST_LINE_REF=:PARM1 AND" _
                                        & " O2.ORDR_LINE_SOURCE=:PARM2"
                    Dim dt As DataTable = oc.GetDataTable(itemSql, New Object() {item.id, orderLineSource})

                    If dt.Rows.Count > 0 Then
                        Dim infoRow As DataRow = dt.Rows(0)

                        Select Case infoRow.Item("ORDR_LINE_STATUS")

                            Case "P", "O", "B"
                                responseItem.OrderStatus = "Open"
                            Case "V", "C", "F"
                                responseItem.OrderStatus = "Closed"
                            Case Else
                                responseItem.OrderStatus = "Open"

                        End Select

                        responseItem.Quantity.Shipped = infoRow.Item("ORDR_QTY_SHIP")
                        responseItem.Quantity.BackOrdered = infoRow.Item("ORDR_QTY_BACK")
                        responseItem.Quantity.Cancelled = infoRow.Item("ORDR_QTY_CANC")


                        responseItem.Shipping.Cost = infoRow.Item("INV_FREIGHT")
                        If infoRow.Item("SHIP_VIA_CODE") & "" <> "" Then
                            Dim trackingURL As String = oc.GetDataValue("Select CARRIER_URL_TRACKING " _
                            & " from SOTCARR1,SOTROUT1,SOTSVIA1 " _
                            & " where SOTSVIA1.SHIP_VIA_CODE = :PARM1" _
                            & "   and SOTROUT1.ROUTE_CODE = SOTSVIA1.ROUTE_CODE " _
                            & "   and SOTCARR1.CARRIER_CODE = SOTROUT1.CARRIER_CODE", New Object() {infoRow.Item("SHIP_VIA_CODE")}) & ""

                            If trackingURL <> "" Then
                                trackingURL = trackingURL & infoRow.Item("SHIP_REF")
                                responseItem.Shipping.TrackingUrl = trackingURL
                            End If
                        End If

                        responseItem.Shipping.Invoice.id = infoRow.Item("INV_NO") & ""
                        If infoRow.Item("INV_DATE") IsNot DBNull.Value Then
                            responseItem.Shipping.ShipDate = infoRow.Item("INV_DATE")
                        End If
                        responseItem.UnitPrice = infoRow.Item("ORDR_UNIT_PRICE")
                        responseItem.StatusDate = DateTime.Parse(infoRow.Item("LAST_DATE")).ToUniversalTime
                        responseItem.Reason = ""
                    Else
                        'check XSTORDR1
                        Dim checkSql As String = "SELECT COUNT(*) FROM XSTORDR2 WHERE ITEM_ID=:PARM1 AND ORDR_SOURCE=:PARM2"

                        Dim existsItem As Integer = oc.GetDataValue(checkSql, New Object() {item.id, statusRequest.Source.type})

                        If existsItem > 0 Then
                            responseItem.OrderStatus = "Acknowledged"
                            responseItem.StatusDate = DateTime.UtcNow

                            responseItem.Quantity.Shipped = 0
                            responseItem.Quantity.BackOrdered = 0
                            responseItem.Quantity.Cancelled = 0
                        Else

                        End If
                    End If

                    statusResponse.Items.itemList.Add(responseItem)
                Next

            Else
                statusResponse.type = "Error"
            End If
        Else
            statusResponse.type = "Error"
        End If

        Return statusResponse
    End Function

#End Region

#Region "Validation"

    Private Function PerformRequestValidation(ByVal orderRequest As OrderRequest, ByVal passPhrase As String) As OrderResponse
        Dim orderResponse As New OrderResponse()

        orderResponse.errorsOccurred = False

        If Not IsPassPhraseValid(passPhrase) Then
            orderResponse.AddRequestError("Invalid passphrase provided")
            Return orderResponse
        End If

        If orderRequest.type <> "Purchase" Then
            orderResponse.AddRequestError(String.Format("Invalid request type ({0})", orderRequest.type))
        End If

        If orderRequest.Source Is Nothing Then
            orderResponse.AddRequestError("Missing Request Source")
        Else
            If Not validSourceTypes.Contains(orderRequest.Source.type) Then
                orderResponse.AddRequestError(String.Format("Invalid Source Type ({0})", orderRequest.Source.type))
            End If
            If orderRequest.Source.Orders Is Nothing OrElse orderRequest.Source.Orders.Count = 0 Then
                orderResponse.AddRequestError("No Orders Provided")
            End If
        End If

        Return orderResponse
    End Function

    Private Function PerformOrderValidation(ByVal order As Order, ByVal orderSource As String) As OrderResult
        Dim orderResult = New OrderResult()
        orderResult.orderID = order.id
        orderResult.orderHasErrors = False

        If Not ValidateOrderID(order.id, orderSource, orderResult) Then
            Return orderResult
        End If

        ValidateCustomerID(order.ClientID, orderResult)
        ValidateOffice(order.Office, orderResult)
        ValidateShipping(order.Shipping, orderSource, orderResult)

        If order.Items Is Nothing OrElse order.Items.itemList.Count = 0 Then
            orderResult.AddOrderError(String.Format("Order {0}: No Items Provided", order.id))
        Else
            For Each item As OrderRequestItem In order.Items.itemList
                ValidateItem(item, orderResult)
            Next
        End If

        Return orderResult
    End Function

    Private Function IsPassPhraseValid(ByVal passPhrase As String) As Boolean
        Return (passPhrase = servicePassPhrase)
    End Function

    Private Function ValidateOrderID(ByVal orderID As String, ByVal orderSource As String, ByRef orderResult As OrderResult) As Boolean
        Dim orderExists As Integer = Val(oc.GetDataValue("SELECT COUNT(*) FROM XSTORDR1 WHERE ORDER_ID=:PARM1 AND ORDR_SOURCE=:PARM2", New Object() {orderID, orderSource}))
        If orderExists > 0 Then
            orderResult.AddOrderError(String.Format("Duplicate Order ID ({0})", orderID))
            Return False
        End If
        Return True
    End Function

    Private Sub ValidateCustomerID(ByVal customerID As String, ByRef orderResult As OrderResult)

        Dim isValid As Integer = 0
        If Not String.IsNullOrEmpty(customerID) Then
            isValid = oc.GetDataValue("SELECT COUNT(*) FROM ARTCUST1 WHERE CUST_CODE=:PARM1", _
                                                     New Object() {customerID.PadLeft(6, "0")})
        End If

        If isValid = 0 Then
            orderResult.AddOrderError(String.Format("Invalid Customer ID ({0})", customerID))
        End If
    End Sub

    Private Sub ValidateOffice(ByVal office As Office, ByRef orderResult As OrderResult)

        'orderResult.AddOrderError(GetFieldError(office.Name, "Office Name", 35))
        orderResult.AddOrderError(GetFieldError(office.Telephone, "Office Telephone", 10))

    End Sub

    Private Sub ValidateShipping(ByVal shipping As Shipping, ByVal orderSource As String, ByRef orderResult As OrderResult)

        ValidateShippingMethod(shipping.Method, orderSource, orderResult)

        If String.IsNullOrEmpty(shipping.Name) Then
            orderResult.AddOrderError("Missing ShipTo Name")
        End If

        orderResult.AddOrderError(GetFieldError(shipping.Name, "ShipTo Name", 35))

        orderResult.AddOrderError(GetFieldError(shipping.Telephone, "ShipTo Telephone", 10))

        If shipping.TaxShipping <> "Y" And shipping.TaxShipping <> "N" Then
            orderResult.AddOrderError(String.Format("Invalid value for TaxShipping ({0})", shipping.TaxShipping))
        End If

        ValidateAddress("ShipTo", shipping.Address, orderResult)

    End Sub

    Private Sub ValidateShippingMethod(ByVal method As String, ByVal orderSource As String, ByRef orderResult As OrderResult)
        Dim isValid As Integer = oc.GetDataValue("SELECT COUNT(*) FROM SOTSVIAF WHERE ORDR_SOURCE=:PARM1 AND SHIP_VIA_DESC=:PARM2", _
                                         New Object() {orderSources(orderSource), method})
        If isValid = 0 Then
            orderResult.AddOrderError(String.Format("Invalid Shipping Method ({0})", method))
        End If
    End Sub


    Private Sub ValidateAddress(ByVal addressType As String, ByVal address As Address, ByRef orderResult As OrderResult)

        orderResult.AddOrderError(GetFieldError(address.AddressLine1, addressType & " Address Line 1", 35))
        orderResult.AddOrderError(GetFieldError(address.City, addressType & " City", 25))
        orderResult.AddOrderError(GetFieldError(address.State, addressType & " State", 2))
        orderResult.AddOrderError(GetFieldError(address.Zip, addressType & " Zip", 10))

        If Not String.IsNullOrEmpty(address.AddressLine2) AndAlso address.AddressLine2.Length > 35 Then
            orderResult.AddOrderError(String.Format("{0} Address Line 2 exceeds maximum length of 35", addressType))
        End If

    End Sub

    Private Function GetFieldError(ByVal field As String, ByVal fieldName As String, ByVal fieldLength As Integer)
        If String.IsNullOrEmpty(field) Then
            Return String.Format("Missing {0} ", fieldName)
        ElseIf field.Length > fieldLength Then
            Return String.Format("{0} exceeds maximum length of {1}", fieldName, fieldLength.ToString)
        End If

        Return ""
    End Function

    Private Sub ValidateItem(ByVal item As OrderRequestItem, ByRef orderResult As OrderResult)
        If String.IsNullOrEmpty(item.Eye) OrElse (item.Eye <> "OD" And item.Eye <> "OS") Then
            orderResult.AddOrderError(String.Format("Invalid/missing L/R Indicator ({0})", item.id))
        End If
        If String.IsNullOrEmpty(item.Product.upc) Then
            Dim prx = item.Product.ProductRx

            If String.IsNullOrEmpty(prx.SER_ID) Then
                orderResult.AddOrderError(String.Format("Item missing SER_ID ({0})", item.id))
            Else
                If String.IsNullOrEmpty(prx.PRD_ADDITION) Then prx.PRD_ADDITION = "0.00"
                Dim addPower As Integer = 0
                If Integer.TryParse(prx.PRD_ADDITION, addPower) Then
                    prx.PRD_ADDITION = addPower.ToString("0.00")
                End If

                If String.IsNullOrEmpty(prx.SER_ID) Then prx.SER_ID = ""
                If String.IsNullOrEmpty(prx.PRD_COLOR) Then prx.PRD_COLOR = ""

                Dim colorCode As String = oc.GetDataValue("SELECT COLOR_CODE FROM ICTCOLR1 WHERE PRICE_CATGY_CODE=:PARM1 AND COLOR_DESC=:PARM2", New Object() {prx.SER_ID, prx.PRD_COLOR}) & ""

                'Stored function checks item master and catalog for this item
                'If it is only in the catalog, it moves it into the item master
                item.itemCode = oc.ExecuteSF("ICPCATLE", _
                                 New String() {"PRICE_CATGY_CODE", "ITEM_BASE_CURVE", _
                                               "ITEM_DIAMETER", "ITEM_SPHERE_POWER", _
                                               "ITEM_CYLINDER", "ITEM_AXIS", _
                                               "ITEM_ADD_POWER", "ITEM_COLOR", "INIT_OPER"}, _
                                 New Object() {prx.SER_ID, prx.PRF_BASECURVE, _
                                               prx.PRF_DIAMETER, prx.PRD_POWER, _
                                               prx.PRD_CYLINDER, prx.PRD_AXIS, _
                                               prx.PRD_ADDITION, colorCode, "webserv"}).ToString

                If String.IsNullOrEmpty(item.itemCode) Then
                    orderResult.AddOrderError(String.Format("Invalid Item ({0})", item.id))
                End If
            End If
        Else
            item.itemCode = oc.ExecuteSF("ICPCATLU", _
                                 New String() {"ITEM_UPC_CODE", "INIT_OPER"}, _
                                 New Object() {item.Product.upc, "webserv"}).ToString

            If String.IsNullOrEmpty(item.itemCode) Then
                orderResult.AddOrderError(String.Format("Invalid UPC ({0})", item.Product.upc))
            End If
        End If
    End Sub

#End Region

#Region "Database Interaction"

    Private Sub WriteOrderToDB(ByVal order As Order, ByVal orderSource As String)
        Try
            oc.BeginTrans()
            Dim docSeqNo As String = oc.ExecuteSF("TAPSEQN1", New String() {"CTL_NO_TYPE_IN"}, New Object() {"XSTORDR1"})
            WriteOrderHeaderToDB(order, docSeqNo, orderSource)
            Dim item_lno As Integer = 1
            For Each item As OrderRequestItem In order.Items.itemList
                WriteOrderItemToDB(item, docSeqNo, item_lno, orderSource)
                item_lno += 1
            Next
            oc.Commit()

        Catch ex As Exception
            oc.Rollback()
            Throw ex
        End Try
    End Sub

    Private Sub WriteOrderHeaderToDB(ByVal order As Order, ByVal docSeqNo As String, ByVal orderSource As String)
        Dim sqlInsertXSTORDR1 = _
        "INSERT INTO XSTORDR1 (XS_DOC_SEQ_NO,ORDER_ID,CUSTOMER_ID, " & _
        "OFFICE_NAME, OFFICE_PHONE, OFFICE_MANAGER, PO_NO, TAX_SHIPPING, " & _
        "SHIPPING_METHOD, SHIP_TO_PATIENT, SHIP_TO_NAME, SHIP_TO_PHONE, SHIP_TO_ADDRESS1, SHIP_TO_ADDRESS2, " & _
        "SHIP_TO_CITY, SHIP_TO_STATE, SHIP_TO_ZIP, PROCESS_IND, TRANSMIT_DATE, " & _
        "PATIENT_ID, PATIENT_DISCOUNT_AMOUNT, PATIENT_SHIPPING_AMOUNT, PATIENT_STAX_RATE, PROMO_CODE, " & _
        "PATIENT_ORDR_AMOUNT, ORDR_SOURCE, ORDR_ORIGIN) " & _
        "VALUES (:PARM1, :PARM2, :PARM3, :PARM4, :PARM5, :PARM6, :PARM7, :PARM8, :PARM9, " & _
        ":PARM10, :PARM11, :PARM12, :PARM13, :PARM14, :PARM15, :PARM16, :PARM17, :PARM18, " & _
        ":PARM19, :PARM20, :PARM21, :PARM22, :PARM23, :PARM24, :PARM25, :PARM26, :PARM27)"


        oc.ExecuteSQL(sqlInsertXSTORDR1, New Object() { _
                      docSeqNo, order.id, order.ClientID, order.Office.Name.Left(35), _
                      order.Office.Telephone, order.Office.Manager, order.PoNumber, If(order.Shipping.TaxShipping = "Y", "1", "0"), _
                      order.Shipping.Method, order.Shipping.ShipToPatient, _
                      order.Shipping.Name.Left(35), order.Shipping.Telephone, order.Shipping.Address.AddressLine1.Left(35), order.Shipping.Address.AddressLine2.Left(35), _
                      order.Shipping.Address.City.Left(25), order.Shipping.Address.State, order.Shipping.Address.Zip, "0", Now.Date, _
                      order.PatientID, order.PatientDiscount, order.PatientShipping, order.PatientStaxRate, order.PromoCode, _
                      order.PatientOrderAmount, orderSource, order.Origin})

    End Sub

    Private Sub WriteOrderItemToDB(ByVal item As OrderRequestItem, ByVal docSeqNo As String, ByVal item_lno As Integer, ByVal orderSource As String)
        Dim sqlInsertXSTORDR2 = _
        "INSERT INTO XSTORDR2 (XS_DOC_SEQ_NO,XS_DOC_SEQ_LNO,ITEM_ID, ITEM_CODE, ORDER_QTY, " & _
        "PATIENT_PRICE, UPC_CODE, ITEM_EYE, PRODUCT_KEY, " & _
        "ITEM_BASE_CURVE, ITEM_DIAMETER, ITEM_SPHERE_POWER, ITEM_CYLINDER, " & _
        "ITEM_AXIS,ITEM_ADD_POWER,ITEM_COLOR,ITEM_MULTIFOCAL,ITEM_NOTE,PATIENT_NAME,ORDR_SOURCE) " & _
        "VALUES (:PARM1, :PARM2, :PARM3, :PARM4, :PARM5, :PARM6, :PARM7, :PARM8, :PARM9, " & _
        ":PARM10, :PARM11, :PARM12, :PARM13, :PARM14, :PARM15, :PARM16, :PARM17, :PARM18, " & _
        ":PARM19,:PARM20)"

        oc.ExecuteSQL(sqlInsertXSTORDR2, New Object() { _
                      docSeqNo, item_lno, item.id, item.itemCode, item.Quantity, _
                      item.PatientPrice, item.Product.upc.Left(12), item.Eye, item.Product.ProductRx.SER_ID, _
                      item.Product.ProductRx.PRF_BASECURVE, item.Product.ProductRx.PRF_DIAMETER, item.Product.ProductRx.PRD_POWER, item.Product.ProductRx.PRD_CYLINDER, _
                      item.Product.ProductRx.PRD_AXIS, item.Product.ProductRx.PRD_ADDITION, item.Product.ProductRx.PRD_COLOR, "", "", item.Patient.Left(40), orderSource})

    End Sub

#End Region

#Region "XML Order Files"

    Private Sub SaveXML(ByVal orderRequest As Order, ByVal orderResult As OrderResult)
        SaveOrderXML(orderRequest)
        SaveOrderResult(orderResult)
    End Sub

    Private Sub SaveOrderXML(ByVal orderRequest As Order)
        Dim xmlStream As New System.IO.FileStream(String.Format(logDirectory & "{0}_req.xml", orderRequest.id), IO.FileMode.Create)
        Dim serializer As XmlSerializer = New XmlSerializer(GetType(Order), xmlns)

        serializer.Serialize(xmlStream, orderRequest)
        xmlStream.Close()
    End Sub

    Private Sub SaveOrderResult(ByVal orderResult As OrderResult)
        Dim xmlStream As New System.IO.FileStream(String.Format(logDirectory & "{0}_res.xml", orderResult.orderID), IO.FileMode.Create)
        Dim serializer As XmlSerializer = New XmlSerializer(GetType(OrderResult), xmlns)

        serializer.Serialize(xmlStream, orderResult)
        xmlStream.Close()
    End Sub

    Private Sub ArchiveOrderXML(ByVal orderID As String)
        System.IO.File.Copy(String.Format(logDirectory & "{0}_req.xml", orderID), String.Format(logDirectory & "XML Archive\{0}_req.xml", orderID), True)
        System.IO.File.Delete(String.Format(logDirectory & "{0}_req.xml", orderID))
        System.IO.File.Delete(String.Format(logDirectory & "{0}_res.xml", orderID))
    End Sub

    Private Sub MoveXMLtoError(ByVal orderID As String, ByVal errorType As ErrorType)

        System.IO.File.Copy(String.Format(logDirectory & "{0}_req.xml", orderID), String.Format(logDirectory & "{0}Errors\{1}_req.xml", errorType.ToString(), orderID), True)
        System.IO.File.Delete(String.Format(logDirectory & "{0}_req.xml", orderID))
        System.IO.File.Copy(String.Format(logDirectory & "{0}_res.xml", orderID), String.Format(logDirectory & "{0}Errors\{1}_res.xml", errorType.ToString(), orderID), True)
        System.IO.File.Delete(String.Format(logDirectory & "{0}_res.xml", orderID))
    End Sub

#End Region

#Region "Error Handling"

    Private Sub LogError(ByVal orderID As String, ByVal ex As Exception)
        'move XML into Error folder
        MoveXMLtoError(orderID, ErrorType.Exception)
        Dim filename As String = String.Format(logDirectory & "{0}Errors\{1}_err.txt", ErrorType.Exception.ToString(), orderID)
        Dim errorFile As New System.IO.FileStream(filename, IO.FileMode.Create)
        Dim errorFileStream As New System.IO.StreamWriter(errorFile)
        errorFileStream.WriteLine(ex.Message)
        errorFileStream.WriteLine(ex.StackTrace)
        errorFileStream.Close()
        errorFile.Close()
    End Sub

#End Region

#Region "Enums"

    Enum ErrorType
        Validation
        Exception
    End Enum

#End Region

End Class
