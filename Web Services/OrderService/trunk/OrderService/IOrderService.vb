Imports System.Runtime.Serialization
Imports System.Xml.Serialization

' NOTE: If you change the interface name "IService1" here, you must also update the reference to "IService1" in App.config.
<ServiceContract(Namespace:="http://www.absolution.com/schemas/ODG/Services"), XmlSerializerFormat()> _
Public Interface IOrderService
    ', ByVal passPhrase As String
    <OperationContract(Name:="PlaceOrder")> _
    Function PlaceOrder(ByVal Request As OrderRequest, ByVal passPhrase As String) As OrderResponse

    <OperationContract(Name:="GetOrderStatus")> _
    Function GetOrderStatus(ByVal Request As StatusRequest, ByVal passPhrase As String) As StatusResponse

End Interface

#Region "StatusRequestClasses"

<XmlRoot(Namespace:="http://www.absolution.com/schemas/ODG/Services")> _
Partial Public Class StatusRequest

    Private typeField As String
    Private sourceField As StatusSource

    Public Sub New()
        MyBase.New()
        sourceField = New StatusSource()
    End Sub

    <XmlAttribute()> _
    Public Property type() As String
        Get
            Return Me.typeField
        End Get
        Set(ByVal value As String)
            Me.typeField = value
        End Set
    End Property

    <XmlElement(IsNullable:=True)> _
    Public Property Source() As StatusSource
        Get
            Return Me.sourceField
        End Get
        Set(ByVal value As StatusSource)
            Me.sourceField = value
        End Set
    End Property
End Class

Partial Public Class StatusSource

    Private typeField As String
    Private itemsField As StatusRequestItems

    Public Sub New()
        MyBase.New()
        itemsField = New StatusRequestItems()
    End Sub

    <XmlAttribute()> _
    Public Property type() As String
        Get
            Return Me.typeField
        End Get
        Set(ByVal value As String)
            Me.typeField = value
        End Set
    End Property

    <XmlElement()> _
    Public Property Items() As StatusRequestItems
        Get
            Return Me.itemsField
        End Get
        Set(ByVal value As StatusRequestItems)
            Me.itemsField = value
        End Set
    End Property
End Class

Public Class StatusRequestItems
    Private itemListField As StatusRequestItemList

    <XmlElement(ElementName:="Item")> _
    Public Property itemList() As StatusRequestItemList
        Get
            Return itemListField
        End Get
        Set(ByVal value As StatusRequestItemList)
            itemListField = value
        End Set
    End Property

    Public Sub New()
        MyBase.New()
        itemListField = New StatusRequestItemList
    End Sub

End Class

<CollectionDataContract(Name:="Items", ItemName:="Item")> _
Public Class StatusRequestItemList
    Inherits List(Of StatusRequestItem)
    Public Sub New()
        MyBase.New()
    End Sub

    Public Sub New(ByVal items() As StatusRequestItem)
        MyBase.New()
        For Each item As StatusRequestItem In items
            Add(item)
        Next item
    End Sub
End Class

Partial Public Class StatusRequestItem

    Private idField As String

    Public Sub New()
        MyBase.New()
    End Sub

    <XmlAttribute()> _
    Public Property id() As String
        Get
            Return Me.idField
        End Get
        Set(ByVal value As String)
            Me.idField = value
        End Set
    End Property
End Class

#End Region

#Region "StatusResponseClasses"

<XmlRoot(Namespace:="http://www.absolution.com/schemas/ODG/Services")> _
Partial Public Class StatusResponse

    Private typeField As String
    Private itemsField As StatusResponseItems

    Public Sub New()
        MyBase.New()
        itemsField = New StatusResponseItems()
    End Sub

    <XmlAttribute()> _
    Public Property type() As String
        Get
            Return Me.typeField
        End Get
        Set(ByVal value As String)
            Me.typeField = value
        End Set
    End Property

    <XmlElement()> _
    Public Property Items() As StatusResponseItems
        Get
            Return Me.itemsField
        End Get
        Set(ByVal value As StatusResponseItems)
            Me.itemsField = value
        End Set
    End Property

End Class


Public Class StatusResponseItems
    Private itemListField As StatusResponseItemList

    <XmlElement(ElementName:="Item")> _
    Public Property itemList() As StatusResponseItemList
        Get
            Return itemListField
        End Get
        Set(ByVal value As StatusResponseItemList)
            itemListField = value
        End Set
    End Property

    Public Sub New()
        MyBase.New()
        itemListField = New StatusResponseItemList
    End Sub

End Class

<CollectionDataContract(Name:="Items", ItemName:="Item")> _
Public Class StatusResponseItemList
    Inherits List(Of StatusResponseItem)
    Public Sub New()
        MyBase.New()
    End Sub

    Public Sub New(ByVal items() As StatusResponseItem)
        MyBase.New()
        For Each item As StatusResponseItem In items
            Add(item)
        Next item
    End Sub
End Class

Partial Public Class StatusResponseItem

    Private idField As String
    Private orderStatusField As String
    Private statusShippingField As StatusShipping
    Private statusQuantityField As StatusQuantity
    Private unitPriceField As Decimal
    Private reasonField As String
    Private statusDateField As DateTime

    Public Sub New()
        MyBase.New()
        statusShippingField = New StatusShipping
        statusQuantityField = New StatusQuantity
    End Sub

    <XmlAttribute()> _
    Public Property id() As String
        Get
            Return Me.idField
        End Get
        Set(ByVal value As String)
            Me.idField = value
        End Set
    End Property

    <XmlElement()> _
    Public Property OrderStatus() As String
        Get
            Return orderStatusField
        End Get
        Set(ByVal value As String)
            orderStatusField = value
        End Set
    End Property

    <XmlElement()> _
    Public Property Shipping() As StatusShipping
        Get
            Return statusShippingField
        End Get
        Set(ByVal value As StatusShipping)
            statusShippingField = value
        End Set
    End Property

    <XmlElement()> _
    Public Property Quantity() As StatusQuantity
        Get
            Return statusQuantityField
        End Get
        Set(ByVal value As StatusQuantity)
            statusQuantityField = value
        End Set
    End Property

    <XmlElement()> _
    Public Property UnitPrice() As Decimal
        Get
            Return unitPriceField
        End Get
        Set(ByVal value As Decimal)
            unitPriceField = value
        End Set
    End Property

    <XmlElement()> _
    Public Property Reason() As String
        Get
            Return reasonField
        End Get
        Set(ByVal value As String)
            reasonField = value
        End Set
    End Property

    <XmlElement()> _
    Public Property StatusDate() As DateTime
        Get
            Return statusDateField
        End Get
        Set(ByVal value As DateTime)
            statusDateField = value
        End Set
    End Property
End Class

Partial Public Class StatusShipping

    Private invoiceField As StatusInvoice
    Private methodField As String
    Private trackingUrlField As String
    Private taxField As Decimal
    Private costField As Decimal
    Private shipDateField As DateTime

    Public Sub New()
        MyBase.New()
        invoiceField = New StatusInvoice
    End Sub

    Public Property Invoice() As StatusInvoice
        Get
            Return invoiceField
        End Get
        Set(ByVal value As StatusInvoice)
            invoiceField = value
        End Set
    End Property

    Public Property Method() As String
        Get
            Return methodField
        End Get
        Set(ByVal value As String)
            methodField = value
        End Set
    End Property

    Public Property TrackingUrl() As String
        Get
            Return trackingUrlField
        End Get
        Set(ByVal value As String)
            trackingUrlField = value
        End Set
    End Property

    Public Property Tax() As Decimal
        Get
            Return taxField
        End Get
        Set(ByVal value As Decimal)
            taxField = value
        End Set
    End Property

    Public Property Cost() As Decimal
        Get
            Return costField
        End Get
        Set(ByVal value As Decimal)
            costField = value
        End Set
    End Property

    Public Property ShipDate() As DateTime
        Get
            Return shipDateField
        End Get
        Set(ByVal value As DateTime)
            shipDateField = value
        End Set
    End Property
End Class

Partial Public Class StatusInvoice
    Private idField As String

    Public Sub New()
        MyBase.New()
    End Sub

    <XmlAttribute()> _
    Public Property id() As String
        Get
            Return idField
        End Get
        Set(ByVal value As String)
            idField = value
        End Set
    End Property
End Class

Partial Public Class StatusQuantity
    Private shippedField As Integer
    Private backOrderedField As Integer
    Private cancelledField As Integer

    Public Sub New()
        MyBase.New()
    End Sub

    Public Property Shipped() As Integer
        Get
            Return shippedField
        End Get
        Set(ByVal value As Integer)
            shippedField = value
        End Set
    End Property

    Public Property BackOrdered() As Integer
        Get
            Return backOrderedField
        End Get
        Set(ByVal value As Integer)
            backOrderedField = value
        End Set
    End Property

    Public Property Cancelled() As Integer
        Get
            Return cancelledField
        End Get
        Set(ByVal value As Integer)
            cancelledField = value
        End Set
    End Property

End Class

#End Region

#Region "OrderRequestClasses"

<XmlRoot(Namespace:="http://www.absolution.com/schemas/ODG/Services", elementName:="Request")> _
Partial Public Class OrderRequest

    Private sourceField As OrderRequestSource
    Private typeField As String

    Public Sub New()
        MyBase.New()
        Me.sourceField = New OrderRequestSource
    End Sub

    <XmlElement(IsNullable:=True)> _
    Public Property Source() As OrderRequestSource
        Get
            Return Me.sourceField
        End Get
        Set(ByVal value As OrderRequestSource)
            Me.sourceField = value
        End Set
    End Property

    <XmlAttribute()> _
    Public Property type() As String
        Get
            Return Me.typeField
        End Get
        Set(ByVal value As String)
            Me.typeField = value
        End Set
    End Property
End Class

Partial Public Class OrderRequestSource

    Private ordersField As List(Of Order)
    Private typeField As String

    Public Sub New()
        MyBase.New()
        Me.ordersField = New List(Of Order)
    End Sub

    <XmlElement(ElementName:="Order")> _
    Public Property Orders() As List(Of Order)
        Get
            Return Me.ordersField
        End Get
        Set(ByVal value As List(Of Order))
            Me.ordersField = value
        End Set
    End Property

    <XmlAttribute()> _
    Public Property type() As String
        Get
            Return Me.typeField
        End Get
        Set(ByVal value As String)
            Me.typeField = value
        End Set
    End Property
End Class

Partial Public Class Order

    Private idField As String
    Private clientIDField As String
    Private officeField As Office
    Private poNumberField As String
    Private shippingField As Shipping
    Private patientIDField As String
    Private patientDiscountField As Decimal?
    Private patientShippingField As Decimal?
    Private patientOrderAmountField As Decimal
    Private patientStaxRateField As Decimal
    Private promoCodeField As String
    Private itemsField As OrderRequestItems
    Private originField As String

    <XmlIgnore()> _
    Public PatientDiscountSpecified As Boolean
    <XmlIgnore()> _
    Public PatientShippingSpecified As Boolean
    <XmlIgnore()> _
    Public PromoCodeSpecified As Boolean


    Public Sub New()
        MyBase.New()
        Me.officeField = New Office
        Me.shippingField = New Shipping
        Me.itemsField = New OrderRequestItems
    End Sub

    <XmlAttribute()> _
    Public Property id() As String
        Get
            Return Me.idField
        End Get
        Set(ByVal value As String)
            Me.idField = value
        End Set
    End Property

    <XmlElement(IsNullable:=True)> _
    Public Property ClientID() As String
        Get
            Return Me.clientIDField
        End Get
        Set(ByVal value As String)
            Me.clientIDField = value
        End Set
    End Property

    <XmlElement(IsNullable:=True)> _
    Public Property Office() As Office
        Get
            Return Me.officeField
        End Get
        Set(ByVal value As Office)
            Me.officeField = value
        End Set
    End Property

    <XmlElement()> _
    Public Property PoNumber() As String
        Get
            Return poNumberField
        End Get
        Set(ByVal value As String)
            poNumberField = value
        End Set
    End Property

    <XmlElement(IsNullable:=True)> _
    Public Property Shipping() As Shipping
        Get
            Return Me.shippingField
        End Get
        Set(ByVal value As Shipping)
            Me.shippingField = value
        End Set
    End Property

    <XmlElement()> _
    Public Property PatientID() As String
        Get
            Return patientIDField
        End Get
        Set(ByVal value As String)
            patientIDField = value
        End Set
    End Property

    <XmlElement()> _
    Public Property PatientDiscount() As Decimal?
        Get
            Return patientDiscountField
        End Get
        Set(ByVal value As Decimal?)
            patientDiscountField = value
            PatientDiscountSpecified = True
        End Set
    End Property

    <XmlElement()> _
    Public Property PatientShipping() As Decimal?
        Get
            Return patientShippingField
        End Get
        Set(ByVal value As Decimal?)
            patientShippingField = value
            PatientShippingSpecified = True
        End Set
    End Property

    <XmlElement()> _
    Public Property PatientOrderAmount() As Decimal
        Get
            Return patientOrderAmountField
        End Get
        Set(ByVal value As Decimal)
            patientOrderAmountField = value
        End Set
    End Property

    <XmlElement()> _
    Public Property PatientStaxRate() As Decimal
        Get
            Return patientStaxRateField
        End Get
        Set(ByVal value As Decimal)
            patientStaxRateField = value
        End Set
    End Property

    <XmlElement()> _
    Public Property PromoCode() As String
        Get
            Return promoCodeField
        End Get
        Set(ByVal value As String)
            promoCodeField = value
            PromoCodeSpecified = True
        End Set
    End Property

    <XmlElement()> _
    Public Property Origin() As String
        Get
            Return Me.originField
        End Get
        Set(ByVal value As String)
            Me.originField = value
        End Set
    End Property


    '<System.Xml.Serialization.XmlArrayAttribute(Order:=3), _
    ' System.Xml.Serialization.XmlArrayItemAttribute("Item", IsNullable:=False)> _
    <XmlElement()> _
    Public Property Items() As OrderRequestItems
        Get
            Return Me.itemsField
        End Get
        Set(ByVal value As OrderRequestItems)
            Me.itemsField = value
        End Set
    End Property

End Class

Partial Public Class Office

    Private nameField As String
    Private telephoneField As String
    Private managerField As String

    Public Sub New()
        MyBase.New()
    End Sub

    <XmlElement(IsNullable:=True)> _
    Public Property Name() As String
        Get
            Return Me.nameField
        End Get
        Set(ByVal value As String)
            Me.nameField = value
        End Set
    End Property

    <XmlElement(IsNullable:=True)> _
    Public Property Telephone() As String
        Get
            Return Me.telephoneField
        End Get
        Set(ByVal value As String)
            Me.telephoneField = value
        End Set
    End Property

    <XmlElement()> _
    Public Property Manager() As String
        Get
            Return Me.managerField
        End Get
        Set(ByVal value As String)
            Me.managerField = value
        End Set
    End Property

End Class

Partial Public Class Shipping

    Private methodField As String
    Private shipToPatientField As String
    Private taxShippingField As String
    Private nameField As String
    Private telephoneField As String
    Private addressField As Address

    Public Sub New()
        MyBase.New()
        Me.addressField = New Address
    End Sub

    <XmlElement(IsNullable:=True)> _
    Public Property Method() As String
        Get
            Return Me.methodField
        End Get
        Set(ByVal value As String)
            Me.methodField = value
        End Set
    End Property

    <XmlElement()> _
    Public Property ShipToPatient() As String
        Get
            Return Me.shipToPatientField
        End Get
        Set(ByVal value As String)
            Me.shipToPatientField = value
        End Set
    End Property

    <XmlElement()> _
    Public Property TaxShipping() As String
        Get
            Return taxShippingField
        End Get
        Set(ByVal value As String)
            taxShippingField = value
        End Set
    End Property

    <XmlElement(IsNullable:=True)> _
    Public Property Name() As String
        Get
            Return Me.nameField
        End Get
        Set(ByVal value As String)
            Me.nameField = value
        End Set
    End Property

    <XmlElement(IsNullable:=True)> _
    Public Property Telephone() As String
        Get
            Return Me.telephoneField
        End Get
        Set(ByVal value As String)
            Me.telephoneField = value
        End Set
    End Property

    <XmlElement(IsNullable:=True)> _
    Public Property Address() As Address
        Get
            Return Me.addressField
        End Get
        Set(ByVal value As Address)
            Me.addressField = value
        End Set
    End Property
End Class

Partial Public Class Address

    Private nameField As String
    Private telephoneField As String
    Private addressLine1Field As String
    Private addressLine2Field As String
    Private cityField As String
    Private stateField As String
    Private zipField As String

    <XmlElement(IsNullable:=True)> _
    Public Property Name() As String
        Get
            Return Me.nameField
        End Get
        Set(ByVal value As String)
            Me.nameField = value
        End Set
    End Property

    <XmlElement(IsNullable:=True)> _
    Public Property Telephone() As String
        Get
            Return Me.telephoneField
        End Get
        Set(ByVal value As String)
            Me.telephoneField = value
        End Set
    End Property

    <XmlElement(IsNullable:=True)> _
    Public Property AddressLine1() As String
        Get
            Return Me.addressLine1Field
        End Get
        Set(ByVal value As String)
            Me.addressLine1Field = value
        End Set
    End Property

    <XmlElement()> _
    Public Property AddressLine2() As String
        Get
            Return Me.addressLine2Field
        End Get
        Set(ByVal value As String)
            Me.addressLine2Field = value
        End Set
    End Property

    <XmlElement(IsNullable:=True)> _
    Public Property City() As String
        Get
            Return Me.cityField
        End Get
        Set(ByVal value As String)
            Me.cityField = value
        End Set
    End Property

    <XmlElement(IsNullable:=True)> _
    Public Property State() As String
        Get
            Return Me.stateField
        End Get
        Set(ByVal value As String)
            Me.stateField = value
        End Set
    End Property

    <XmlElement(IsNullable:=True)> _
    Public Property Zip() As String
        Get
            Return Me.zipField
        End Get
        Set(ByVal value As String)
            Me.zipField = value
        End Set
    End Property
End Class

Public Class OrderRequestItems
    Private itemListField As OrderRequestItemList

    <XmlElement(ElementName:="Item")> _
    Public Property itemList() As OrderRequestItemList
        Get
            Return itemListField
        End Get
        Set(ByVal value As OrderRequestItemList)
            itemListField = value
        End Set
    End Property

    Public Sub New()
        MyBase.New()
        itemListField = New OrderRequestItemList
    End Sub

End Class

<CollectionDataContract(Name:="Items", ItemName:="Item")> _
Public Class OrderRequestItemList
    Inherits List(Of OrderRequestItem)
    Public Sub New()
        MyBase.New()
    End Sub

    Public Sub New(ByVal items() As OrderRequestItem)
        MyBase.New()
        For Each item As OrderRequestItem In items
            Add(item)
        Next item
    End Sub
End Class

Partial Public Class OrderRequestItem

    Private idField As String
    Private patientField As String
    Private eyeField As String
    Private quantityField As UShort
    Private patientPriceField As Decimal
    Private productField As Product
    Private itemCodeField As String

    Public Sub New()
        MyBase.New()
        Me.productField = New Product
    End Sub

    <XmlAttribute()> _
    Public Property id() As String
        Get
            Return Me.idField
        End Get
        Set(ByVal value As String)
            Me.idField = value
        End Set
    End Property

    <XmlIgnore()> _
    Public Property itemCode() As String
        Get
            Return itemCodeField
        End Get
        Set(ByVal value As String)
            itemCodeField = value
        End Set
    End Property

    <XmlElement(IsNullable:=True)> _
    Public Property Patient() As String
        Get
            Return Me.patientField
        End Get
        Set(ByVal value As String)
            Me.patientField = value
        End Set
    End Property

    <XmlElement(IsNullable:=True)> _
    Public Property Eye() As String
        Get
            Return Me.eyeField
        End Get
        Set(ByVal value As String)
            Me.eyeField = value
        End Set
    End Property

    <XmlElement()> _
    Public Property Quantity() As UShort
        Get
            Return Me.quantityField
        End Get
        Set(ByVal value As UShort)
            Me.quantityField = value
        End Set
    End Property

    <XmlElement()> _
    Public Property PatientPrice() As Decimal
        Get
            Return Me.patientPriceField
        End Get
        Set(ByVal value As Decimal)
            Me.patientPriceField = value
        End Set
    End Property

    <XmlElement()> _
    Public Property Product() As Product
        Get
            Return Me.productField
        End Get
        Set(ByVal value As Product)
            Me.productField = value
        End Set
    End Property

End Class

Partial Public Class Product

    Private upcField As String
    Private productIDField As String
    Private productRxField As ProductRX

    Public Sub New()
        MyBase.New()
        productRxField = New ProductRX
    End Sub

    Public Property upc() As String
        Get
            Return Me.upcField
        End Get
        Set(ByVal value As String)
            Me.upcField = value
        End Set
    End Property

    Public Property productID() As String
        Get
            Return Me.productIDField
        End Get
        Set(ByVal value As String)
            Me.productIDField = value
        End Set
    End Property

    Public Property ProductRx() As ProductRX
        Get
            Return Me.productRxField
        End Get
        Set(ByVal value As ProductRX)
            Me.productRxField = value
        End Set
    End Property
End Class

Partial Public Class ProductRX

    Private sER_IDField As String
    Private pRF_BASECURVEField As Decimal
    Private pRF_DIAMETERField As Decimal
    Private pRD_POWERField As Decimal
    Private pRD_CYLINDERField As Decimal?
    Private pRD_AXISField As UShort?
    Private pRD_ADDITIONField As String
    Private pRD_COLORField As String

    <XmlIgnore()> _
    Public PRD_CYLINDERSpecified As Boolean

    <XmlIgnore()> _
    Public PRD_AXISSpecified As Boolean

    <XmlElement(IsNullable:=False)> _
    Public Property SER_ID() As String
        Get
            Return Me.sER_IDField
        End Get
        Set(ByVal value As String)
            Me.sER_IDField = value
        End Set
    End Property

    <XmlElement()> _
    Public Property PRF_BASECURVE() As Decimal
        Get
            Return Me.pRF_BASECURVEField
        End Get
        Set(ByVal value As Decimal)
            Me.pRF_BASECURVEField = value
        End Set
    End Property

    <XmlElement()> _
    Public Property PRF_DIAMETER() As Decimal
        Get
            Return Me.pRF_DIAMETERField
        End Get
        Set(ByVal value As Decimal)
            Me.pRF_DIAMETERField = value
        End Set
    End Property

    <XmlElement()> _
    Public Property PRD_POWER() As Decimal
        Get
            Return Me.pRD_POWERField
        End Get
        Set(ByVal value As Decimal)
            Me.pRD_POWERField = value
        End Set
    End Property

    <XmlElement()> _
    Public Property PRD_CYLINDER() As Decimal?
        Get
            Return Me.pRD_CYLINDERField
        End Get
        Set(ByVal value As Decimal?)
            Me.pRD_CYLINDERField = value
            PRD_CYLINDERSpecified = True
        End Set
    End Property

    <XmlElement()> _
    Public Property PRD_AXIS() As UShort?
        Get
            Return Me.pRD_AXISField
        End Get
        Set(ByVal value As UShort?)
            Me.pRD_AXISField = value
            PRD_AXISSpecified = True
        End Set
    End Property

    <XmlElement()> _
    Public Property PRD_ADDITION() As String
        Get
            Return Me.pRD_ADDITIONField
        End Get
        Set(ByVal value As String)
            Me.pRD_ADDITIONField = value
        End Set
    End Property

    <XmlElement()> _
    Public Property PRD_COLOR() As String
        Get
            Return Me.pRD_COLORField
        End Get
        Set(ByVal value As String)
            Me.pRD_COLORField = value
        End Set
    End Property

    Public Sub New()

    End Sub
End Class

#End Region

#Region "OrderResultClasses"

<XmlRoot()> _
Partial Public Class OrderResponse

    Private orderResultsField As List(Of OrderResult)
    Private errorsOccurredField As Boolean
    Private requestErrorsField As List(Of String)

    <XmlElement()> _
    Public Property errorsOccurred() As Boolean
        Get
            Return Me.errorsOccurredField
        End Get
        Set(ByVal value As Boolean)
            Me.errorsOccurredField = value
        End Set
    End Property

    <XmlElement(ElementName:="OrderResults")> _
    Public Property OrderResults() As List(Of OrderResult)
        Get
            Return Me.orderResultsField
        End Get
        Set(ByVal value As List(Of OrderResult))
            Me.orderResultsField = value
        End Set
    End Property

    <XmlElement(ElementName:="RequestErrors")> _
    Public Property RequestErrors() As List(Of String)
        Get
            Return Me.requestErrorsField
        End Get
        Set(ByVal value As List(Of String))
            Me.requestErrorsField = value
        End Set
    End Property

    Friend Sub AddRequestError(ByVal errorText As String)
        Me.errorsOccurred = True
        Me.RequestErrors.Add(errorText)
    End Sub

    Public Sub New()
        MyBase.New()
        Me.requestErrorsField = New List(Of String)
        Me.orderResultsField = New List(Of OrderResult)
    End Sub
End Class

Partial Public Class OrderResult
    Private orderIDField As String
    Private orderHasErrorsField As Boolean

    Private orderErrorTextField As List(Of String)

    <XmlElement()> _
    Public Property orderID() As String
        Get
            Return Me.orderIDField
        End Get
        Set(ByVal value As String)
            Me.orderIDField = value
        End Set
    End Property

    <XmlElement()> _
    Public Property orderHasErrors() As Boolean
        Get
            Return Me.orderHasErrorsField
        End Get
        Set(ByVal value As Boolean)
            Me.orderHasErrorsField = value
        End Set
    End Property

    <XmlElement(ElementName:="OrderError")> _
    Public Property orderErrors() As List(Of String)
        Get
            Return Me.orderErrorTextField
        End Get
        Set(ByVal value As List(Of String))
            Me.orderErrorTextField = value
        End Set
    End Property

    Public Sub New()
        MyBase.New()
        Me.orderErrorTextField = New List(Of String)
    End Sub

    Friend Sub AddOrderError(ByVal errorText As String)
        If Not String.IsNullOrEmpty(errorText) Then
            Me.orderHasErrors = True
            Me.orderErrors.Add(errorText)
        End If
    End Sub

End Class

#End Region
