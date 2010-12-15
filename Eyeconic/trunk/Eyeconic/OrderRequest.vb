Imports System.Xml.Serialization
' NOTE: If you change the class name "OrderRequest" here, you must also update the reference to "OrderRequest" in App.config.
Public Class OrderRequest
    Implements IOrderRequest

    Private oc As oracleClient

    Public Function PlaceOrder(ByVal orderRequest As Request) As OrderResponse Implements IOrderRequest.PlaceOrder

        Dim orderResponse = PerformRequestValidation(orderRequest)

        If Not orderResponse.errorsOccurred Then
            oc = New oracleClient()

            For Each order As Order In orderRequest.Source.Orders
                SetOrderDefaults(order)
                Dim orderResult As OrderResult = PerformOrderValidation(order)

                SaveXML(order, orderResult)

                If Not orderResult.orderHasErrors Then
                    Try
                        WriteOrderToDB(order, orderRequest.Source.type)
                        ArchiveOrderXML(order.id)
                    Catch ex As Exception
                        LogError(order.id, ex)
                    End Try
                Else
                    MoveXMLtoError(order.id)
                    orderResponse.errorsOccurred = True
                End If

                orderResponse.OrderResults.Add(orderResult)
            Next
        End If

        Return orderResponse
    End Function

    Private Sub SetOrderDefaults(ByVal order As Order)
        If String.IsNullOrEmpty(order.Shipping.TaxShipping) Then
            order.Shipping.TaxShipping = "N"
        End If
    End Sub

#Region "Validation"

    Private Function PerformRequestValidation(ByVal orderRequest As Request) As OrderResponse
        Dim orderResponse As New OrderResponse()

        orderResponse.errorsOccurred = False

        If orderRequest.type <> "Purchase" Then
            orderResponse.AddRequestError(String.Format("Invalid request type ({0})", orderRequest.type))
        End If

        If orderRequest.Source Is Nothing Then
            orderResponse.AddRequestError("Missing Request Source")
        Else
            If orderRequest.Source.type <> "VSP" Then
                orderResponse.AddRequestError(String.Format("Invalid Source Type ({0})", orderRequest.Source.type))
            End If
            If orderRequest.Source.Orders Is Nothing OrElse orderRequest.Source.Orders.Count = 0 Then
                orderResponse.AddRequestError("No Orders Provided")
            End If
        End If

        Return orderResponse
    End Function

    Private Function PerformOrderValidation(ByVal order As Order) As OrderResult
        Dim orderResult = New OrderResult()
        orderResult.orderID = order.id
        orderResult.orderHasErrors = False

        If Not ValidateOrderID(order.id, orderResult) Then
            Return orderResult
        End If

        ValidateCustomerID(order.CustomerID, orderResult)
        ValidateOffice(order.Office, orderResult)
        ValidateShipping(order.Shipping, orderResult)

        If order.Items Is Nothing OrElse order.Items.itemList.Count = 0 Then
            orderResult.AddOrderError(String.Format("Order {0}: No Items Provided", order.id))
        Else
            For Each item As Item In order.Items.itemList
                ValidateItem(item, orderResult)
            Next
        End If

        Return orderResult
    End Function

    Private Function ValidateOrderID(ByVal orderID As String, ByRef orderResult As OrderResult) As Boolean
        Dim orderExists As Integer = Val(oc.GetDataValue("SELECT COUNT(*) FROM XSTORDR1 WHERE ORDER_ID=:PARM1", New Object() {orderID}))
        If orderExists > 0 Then
            orderResult.AddOrderError(String.Format("Duplicate Order ID ({0})", orderID))
            Return False
        End If
        Return True
    End Function

    Private Sub ValidateCustomerID(ByVal customerID As String, ByRef orderResult As OrderResult)
        Dim isValid As Integer = oc.GetDataValue("SELECT COUNT(*) FROM ARTCUST1 WHERE CUST_CODE=:PARM1", _
                                                 New Object() {customerID.PadLeft(6, "0")})
        If isValid = 0 Then
            orderResult.AddOrderError(String.Format("Invalid Customer ID ({0})", customerID))
        End If
    End Sub

    Private Sub ValidateOffice(ByVal office As Office, ByRef orderResult As OrderResult)
        If String.IsNullOrEmpty(office.Name) Then
            orderResult.AddOrderError("Missing Office Name")
        End If

        If String.IsNullOrEmpty(office.Telephone) Then
            orderResult.AddOrderError("Missing Office Telephone")
        End If

        ValidateAddress("Office", office.Address, orderResult)
    End Sub

    Private Sub ValidateShipping(ByVal shipping As Shipping, ByRef orderResult As OrderResult)

        ValidateShippingMethod(shipping.Method, orderResult)

        If String.IsNullOrEmpty(shipping.Name) Then
            orderResult.AddOrderError("Missing ShipTo Name")
        End If

        If String.IsNullOrEmpty(shipping.Telephone) Then
            orderResult.AddOrderError("Missing ShipTo Telephone")
        End If

        If shipping.TaxShipping <> "Y" And shipping.TaxShipping <> "N" Then
            orderResult.AddOrderError(String.Format("Invalid value for TaxShipping ({0})", shipping.TaxShipping))
        End If

        ValidateAddress("ShipTo", shipping.Address, orderResult)

    End Sub

    Private Sub ValidateShippingMethod(ByVal method As String, ByRef orderResult As OrderResult)
        Dim isValid As Integer = oc.GetDataValue("SELECT COUNT(*) FROM SOTSVIAF WHERE ORDR_SOURCE='Y' AND SHIP_VIA_DESC=:PARM1", _
                                         New Object() {method})
        If isValid = 0 Then
            orderResult.AddOrderError(String.Format("Invalid Shipping Method ({0})", method))
        End If
    End Sub

    Private Sub ValidateAddress(ByVal addressType As String, ByVal address As Address, ByRef orderResult As OrderResult)
        If String.IsNullOrEmpty(address.AddressLine1) Then
            orderResult.AddOrderError(String.Format("Missing {0} Address Line 1", addressType))
        End If
        If String.IsNullOrEmpty(address.City) Then
            orderResult.AddOrderError(String.Format("Missing {0} City", addressType))
        End If
        If String.IsNullOrEmpty(address.State) Then
            orderResult.AddOrderError(String.Format("Missing {0} State", addressType))
        End If
        If String.IsNullOrEmpty(address.Zip) Then
            orderResult.AddOrderError(String.Format("Missing {0} Zip", addressType))
        End If
    End Sub

    Private Sub ValidateItem(ByVal item As Item, ByRef orderResult As OrderResult)
        If String.IsNullOrEmpty(item.Eye) OrElse (item.Eye <> "OD" And item.Eye <> "OS") Then
            orderResult.AddOrderError(String.Format("Invalid/missing L/R Indicator ({0})", item.id))
        End If
        If String.IsNullOrEmpty(item.Product.upc) Then
            Dim prx = item.Product.ProductRx

            If String.IsNullOrEmpty(prx.SER_ID) Then
                orderResult.AddOrderError(String.Format("Item missing SER_ID ({0})", item.id))
            Else
                If String.IsNullOrEmpty(prx.PRD_ADDITION) Then prx.PRD_ADDITION = "0.00"

                Dim colorCode As String = oc.GetDataValue("SELECT COLOR_CODE FROM ICTCOLR1 WHERE PRICE_CATGY_CODE=:PARM1 AND COLOR_DESC=:PARM2", New Object() {prx.SER_ID, prx.PRD_COLOR})

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
            'check UPC
        End If
    End Sub

#End Region

#Region "Database Interaction"

    Private Sub WriteOrderToDB(ByVal order As Order, ByVal orderSource As String)
        Try
            oc.BeginTrans()
            Dim docSeqNo As String = oc.ExecuteSF("TAPSEQN1", New String() {"CTL_NO_TYPE_IN"}, New Object() {"XSTORDR1"})
            WriteOrderHeaderToDB(order, docSeqNo)
            Dim item_lno As Integer = 1
            For Each item As Item In order.Items.itemList
                WriteOrderItemToDB(item, docSeqNo, orderSource, item_lno)
                item_lno += 1
            Next
            oc.Commit()
        Catch ex As Exception
            oc.Rollback()
            Throw ex
        End Try
    End Sub

    Private Sub WriteOrderHeaderToDB(ByVal order As Order, ByVal docSeqNo As String)
        Dim sqlInsertXSTORDR1 = _
        "INSERT INTO XSTORDR1 (XS_DOC_SEQ_NO,ORDER_ID,CUSTOMER_ID, OFFICE_ID, OFFICE_NAME, " & _
        "OFFICE_PHONE, OFFICE_SHIP_TO_ADDRESS1, OFFICE_SHIP_TO_ADDRESS2, OFFICE_SHIP_TO_CITY, " & _
        "OFFICE_SHIP_TO_STATE, OFFICE_SHIP_TO_ZIP, PATIENT_NAME, SHIPPING_METHOD, " & _
        "SHIP_TO_PATIENT, SHIP_TO_NAME, SHIP_TO_PHONE, SHIP_TO_ADDRESS1, SHIP_TO_ADDRESS2, " & _
        "SHIP_TO_CITY, SHIP_TO_STATE, SHIP_TO_ZIP, PROCESS_IND, TRANSMIT_DATE, TAX_SHIPPING, PATIENT_STAX_RATE) " & _
        "VALUES (:PARM1, :PARM2, :PARM3, :PARM4, :PARM5, :PARM6, :PARM7, :PARM8, :PARM9, " & _
        ":PARM10, :PARM11, :PARM12, :PARM13, :PARM14, :PARM15, :PARM16, :PARM17, :PARM18, " & _
        ":PARM19, :PARM20, :PARM21, :PARM22, :PARM23, :PARM24, :PARM25)"


        oc.ExecuteSQL(sqlInsertXSTORDR1, New Object() { _
                      docSeqNo, order.id, order.CustomerID, order.Office.OfficeID, order.Office.Name, _
                      order.Office.Telephone, order.Office.Address.AddressLine1, order.Office.Address.AddressLine2 & "", order.Office.Address.City, _
                      order.Office.Address.State, order.Office.Address.Zip, order.Shipping.Name, order.Shipping.Method, _
                      order.Shipping.ShipToPatient, order.Shipping.Name, order.Shipping.Telephone, order.Shipping.Address.AddressLine1, order.Shipping.Address.AddressLine2 & "", _
                      order.Shipping.Address.City, order.Shipping.Address.State, order.Shipping.Address.Zip, "0", Now.Date, If(order.Shipping.TaxShipping = "Y", "1", "0"), order.PatientStaxRate})

    End Sub

    Private Sub WriteOrderItemToDB(ByVal item As Item, ByVal docSeqNo As String, ByVal orderSource As String, ByVal item_lno As Integer)
        Dim sqlInsertXSTORDR2 = _
        "INSERT INTO XSTORDR2 (XS_DOC_SEQ_NO,XS_DOC_SEQ_LNO,ITEM_ID, ITEM_CODE, ORDER_QTY, " & _
        "PATIENT_PRICE, UPC_CODE, ITEM_EYE, PRODUCT_KEY, " & _
        "ITEM_BASE_CURVE, ITEM_DIAMETER, ITEM_SPHERE_POWER, ITEM_CYLINDER, " & _
        "ITEM_AXIS,ITEM_ADD_POWER,ITEM_COLOR,ITEM_MULTIFOCAL,ITEM_NOTE,PATIENT_NAME, ORDR_SOURCE) " & _
        "VALUES (:PARM1, :PARM2, :PARM3, :PARM4, :PARM5, :PARM6, :PARM7, :PARM8, :PARM9, " & _
        ":PARM10, :PARM11, :PARM12, :PARM13, :PARM14, :PARM15, :PARM16, :PARM17, :PARM18, " & _
        ":PARM19, :PARM20)"

        oc.ExecuteSQL(sqlInsertXSTORDR2, New Object() { _
                      docSeqNo, item_lno, item.id, item.itemCode, item.Quantity, _
                      item.PatientPrice, item.Product.upc & "", item.Eye & "", item.Product.ProductRx.SER_ID & "", _
                      item.Product.ProductRx.PRF_BASECURVE, item.Product.ProductRx.PRF_DIAMETER, item.Product.ProductRx.PRD_POWER, item.Product.ProductRx.PRD_CYLINDER, _
                      item.Product.ProductRx.PRD_AXIS, item.Product.ProductRx.PRD_ADDITION & "", item.Product.ProductRx.PRD_COLOR & "", "", "", item.Patient & "", orderSource})

    End Sub

#End Region

#Region "XML Order Files"

    Private Sub SaveXML(ByVal orderRequest As Order, ByVal orderResult As OrderResult)
        SaveOrderXML(orderRequest)
        SaveOrderResult(orderResult)
    End Sub

    Private Sub SaveOrderXML(ByVal orderRequest As Order)
        Dim xmlStream As New System.IO.FileStream(String.Format("C:\Eyeconic\{0}_req.xml", orderRequest.id), IO.FileMode.Create)
        Dim serializer As XmlSerializer = New XmlSerializer(GetType(Order))
        serializer.Serialize(xmlStream, orderRequest)
        xmlStream.Close()
    End Sub

    Private Sub SaveOrderResult(ByVal orderResult As OrderResult)
        Dim xmlStream As New System.IO.FileStream(String.Format("C:\Eyeconic\{0}_res.xml", orderResult.orderID), IO.FileMode.Create)
        Dim serializer As XmlSerializer = New XmlSerializer(GetType(OrderResult))
        serializer.Serialize(xmlStream, orderResult)
        xmlStream.Close()
    End Sub

    Private Sub ArchiveOrderXML(ByVal orderID As String)
        System.IO.File.Copy(String.Format("C:\Eyeconic\{0}_req.xml", orderID), String.Format("C:\Eyeconic\XML Archive\{0}_req.xml", orderID), True)
        System.IO.File.Delete(String.Format("C:\Eyeconic\{0}_req.xml", orderID))
        System.IO.File.Delete(String.Format("C:\Eyeconic\{0}_res.xml", orderID))
    End Sub

    Private Sub MoveXMLtoError(ByVal orderID As String)
        System.IO.File.Copy(String.Format("C:\Eyeconic\{0}_req.xml", orderID), String.Format("C:\Eyeconic\Errors\{0}_req.xml", orderID), True)
        System.IO.File.Delete(String.Format("C:\Eyeconic\{0}_req.xml", orderID))
        System.IO.File.Copy(String.Format("C:\Eyeconic\{0}_res.xml", orderID), String.Format("C:\Eyeconic\Errors\{0}_res.xml", orderID), True)
        System.IO.File.Delete(String.Format("C:\Eyeconic\{0}_res.xml", orderID))
    End Sub

#End Region

#Region "Error Handling"

    Private Sub LogError(ByVal orderID As String, ByVal ex As Exception)
        'move XML into Error folder
        MoveXMLtoError(orderID)
        Dim filename As String = String.Format("C:\Eyeconic\Errors\{0}_err.txt", orderID)
        Dim errorFile As New System.IO.FileStream(filename, IO.FileMode.Create)
        Dim errorFileStream As New System.IO.StreamWriter(errorFile)
        errorFileStream.WriteLine(ex.Message)
        errorFileStream.WriteLine(ex.StackTrace)
        errorFileStream.Close()
        errorFile.Close()
    End Sub

#End Region

End Class
