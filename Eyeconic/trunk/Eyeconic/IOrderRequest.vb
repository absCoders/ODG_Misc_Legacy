Imports System.Runtime.Serialization
Imports System.Xml.Serialization

' NOTE: If you change the interface name "IService1" here, you must also update the reference to "IService1" in App.config.
<ServiceContract(), XmlSerializerFormat()> _
Public Interface IOrderRequest

    <OperationContract(Name:="PlaceOrder")> _
    Function PlaceOrder(ByVal value As Request) As OrderResponse

End Interface


#Region "OrderRequestClasses"

<DataContract()> _
<XmlRoot()> _
Partial Public Class Request

    Private sourceField As Source
    Private typeField As String

    Public Sub New()
        MyBase.New()
        Me.sourceField = New Source
    End Sub

    <DataMember(IsRequired:=True)> _
    <XmlElement(IsNullable:=False)> _
    Public Property Source() As Source
        Get
            Return Me.sourceField
        End Get
        Set(ByVal value As Source)
            Me.sourceField = value
        End Set
    End Property

    <DataMember(IsRequired:=True)> _
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

<DataContract()> _
Partial Public Class Source

    Private ordersField As List(Of Order)

    Private typeField As String

    Public Sub New()
        MyBase.New()
        Me.ordersField = New List(Of Order)
    End Sub

    <DataMember(IsRequired:=True)> _
    <XmlElement(ElementName:="Order")> _
    Public Property Orders() As List(Of Order)
        Get
            Return Me.ordersField
        End Get
        Set(ByVal value As List(Of Order))
            Me.ordersField = value
        End Set
    End Property

    <DataMember(IsRequired:=True)> _
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


<DataContract()> _
Partial Public Class Order

    Private customerIDField As String
    Private officeField As Office
    Private shippingField As Shipping
    Private patientStaxRateField As Decimal
    Private itemsField As Items
    Private idField As String

    Public Sub New()
        MyBase.New()
        Me.itemsField = New Items
        Me.shippingField = New Shipping
        Me.officeField = New Office
    End Sub

    <DataMember(IsRequired:=True)> _
    <XmlElement()> _
    Public Property CustomerID() As String
        Get
            Return Me.customerIDField
        End Get
        Set(ByVal value As String)
            Me.customerIDField = value
        End Set
    End Property

    <DataMember(IsRequired:=True)> _
    <XmlElement()> _
    Public Property Office() As Office
        Get
            Return Me.officeField
        End Get
        Set(ByVal value As Office)
            Me.officeField = value
        End Set
    End Property

    <DataMember(IsRequired:=True)> _
    <XmlElement()> _
    Public Property Shipping() As Shipping
        Get
            Return Me.shippingField
        End Get
        Set(ByVal value As Shipping)
            Me.shippingField = value
        End Set
    End Property

    <DataMember(isrequired:=True)> _
    <XmlElement()> _
    Public Property PatientStaxRate() As Decimal
        Get
            Return patientStaxRateField
        End Get
        Set(ByVal value As Decimal)
            patientStaxRateField = value
        End Set
    End Property

    '<System.Xml.Serialization.XmlArrayAttribute(Order:=3), _
    ' System.Xml.Serialization.XmlArrayItemAttribute("Item", IsNullable:=False)> _
    <DataMember(IsRequired:=True)> _
    <XmlElement()> _
    Public Property Items() As Items
        Get
            Return Me.itemsField
        End Get
        Set(ByVal value As Items)
            Me.itemsField = value
        End Set
    End Property

    <DataMember(IsRequired:=True)> _
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

<DataContract()> _
Partial Public Class Office

    Private officeIDField As UInteger

    Private nameField As String

    Private telephoneField As String

    Private addressField As Address

    Public Sub New()
        MyBase.New()
        Me.addressField = New Address
    End Sub

    <DataMember(IsRequired:=True)> _
    Public Property OfficeID() As UInteger
        Get
            Return Me.officeIDField
        End Get
        Set(ByVal value As UInteger)
            Me.officeIDField = value
        End Set
    End Property

    <DataMember(IsRequired:=True)> _
    Public Property Name() As String
        Get
            Return Me.nameField
        End Get
        Set(ByVal value As String)
            Me.nameField = value
        End Set
    End Property

    <DataMember(IsRequired:=True)> _
    Public Property Telephone() As String
        Get
            Return Me.telephoneField
        End Get
        Set(ByVal value As String)
            Me.telephoneField = value
        End Set
    End Property

    <DataMember(IsRequired:=True)> _
    Public Property Address() As Address
        Get
            Return Me.addressField
        End Get
        Set(ByVal value As Address)
            Me.addressField = value
        End Set
    End Property
End Class

<DataContract()> _
Partial Public Class Shipping

    Private methodField As String
    Private shipToPatientField As String
    Private nameField As String
    Private telephoneField As String
    Private taxShippingField As String
    Private addressField As Address

    Public Sub New()
        MyBase.New()
        Me.addressField = New Address
    End Sub

    <DataMember(IsRequired:=True)> _
    Public Property Method() As String
        Get
            Return Me.methodField
        End Get
        Set(ByVal value As String)
            Me.methodField = value
        End Set
    End Property

    <DataMember(IsRequired:=False)> _
    Public Property TaxShipping() As String
        Get
            Return taxShippingField
        End Get
        Set(ByVal value As String)
            taxShippingField = value
        End Set
    End Property

    <DataMember(IsRequired:=False)> _
    Public Property ShipToPatient() As String
        Get
            Return Me.shipToPatientField
        End Get
        Set(ByVal value As String)
            Me.shipToPatientField = value
        End Set
    End Property

    <DataMember(IsRequired:=True)> _
    Public Property Name() As String
        Get
            Return Me.nameField
        End Get
        Set(ByVal value As String)
            Me.nameField = value
        End Set
    End Property

    <DataMember(IsRequired:=True)> _
    Public Property Telephone() As String
        Get
            Return Me.telephoneField
        End Get
        Set(ByVal value As String)
            Me.telephoneField = value
        End Set
    End Property

    <DataMember(IsRequired:=True)> _
    Public Property Address() As Address
        Get
            Return Me.addressField
        End Get
        Set(ByVal value As Address)
            Me.addressField = value
        End Set
    End Property
End Class

<DataContract()> _
Partial Public Class Address

    Private addressLine1Field As String

    Private addressLine2Field As String

    Private cityField As String

    Private stateField As String

    Private zipField As String

    <DataMember(IsRequired:=True)> _
    Public Property AddressLine1() As String
        Get
            Return Me.addressLine1Field
        End Get
        Set(ByVal value As String)
            Me.addressLine1Field = value
        End Set
    End Property

    <DataMember(IsRequired:=False)> _
    Public Property AddressLine2() As String
        Get
            Return Me.addressLine2Field
        End Get
        Set(ByVal value As String)
            Me.addressLine2Field = value
        End Set
    End Property

    <DataMember(IsRequired:=True)> _
    Public Property City() As String
        Get
            Return Me.cityField
        End Get
        Set(ByVal value As String)
            Me.cityField = value
        End Set
    End Property

    <DataMember(IsRequired:=True)> _
    Public Property State() As String
        Get
            Return Me.stateField
        End Get
        Set(ByVal value As String)
            Me.stateField = value
        End Set
    End Property

    <DataMember(IsRequired:=True)> _
    Public Property Zip() As String
        Get
            Return Me.zipField
        End Get
        Set(ByVal value As String)
            Me.zipField = value
        End Set
    End Property
End Class


Public Class Items
    Private itemListField As ItemList

    <DataMember(IsRequired:=True)> _
    <XmlElement(ElementName:="Item")> _
    Public Property itemList() As ItemList
        Get
            Return itemListField
        End Get
        Set(ByVal value As ItemList)
            itemListField = value
        End Set
    End Property

    Public Sub New()
        MyBase.New()
        itemListField = New ItemList
    End Sub

End Class

<CollectionDataContract(Name:="Items", ItemName:="Item")> _
Public Class ItemList
    Inherits List(Of Item)
    Public Sub New()
        MyBase.New()
    End Sub

    Public Sub New(ByVal items() As Item)
        MyBase.New()
        For Each item As Item In items
            Add(item)
        Next item
    End Sub
End Class

<DataContract()> _
Partial Public Class Item

    Private patientField As String
    Private eyeField As String
    Private quantityField As UShort
    Private patientPriceField As Decimal
    Private productField As Product
    Private itemCodeField As String
    Private idField As String

    Public Sub New()
        MyBase.New()
        Me.productField = New Product
    End Sub

    <XmlIgnore()> _
    Public Property itemCode() As String
        Get
            Return itemCodeField
        End Get
        Set(ByVal value As String)
            itemCodeField = value
        End Set
    End Property

    <DataMember(IsRequired:=True)> _
    <XmlElement()> _
    Public Property Patient() As String
        Get
            Return Me.patientField
        End Get
        Set(ByVal value As String)
            Me.patientField = value
        End Set
    End Property

    <DataMember(IsRequired:=True)> _
    <XmlElement()> _
    Public Property Eye() As String
        Get
            Return Me.eyeField
        End Get
        Set(ByVal value As String)
            Me.eyeField = value
        End Set
    End Property

    <DataMember(IsRequired:=True)> _
    <XmlElement()> _
    Public Property Quantity() As UShort
        Get
            Return Me.quantityField
        End Get
        Set(ByVal value As UShort)
            Me.quantityField = value
        End Set
    End Property

    <DataMember(IsRequired:=True)> _
    <XmlElement()> _
    Public Property PatientPrice() As Decimal
        Get
            Return Me.patientPriceField
        End Get
        Set(ByVal value As Decimal)
            Me.patientPriceField = value
        End Set
    End Property

    <DataMember(IsRequired:=True)> _
    <XmlElement()> _
    Public Property Product() As Product
        Get
            Return Me.productField
        End Get
        Set(ByVal value As Product)
            Me.productField = value
        End Set
    End Property

    <DataMember(IsRequired:=True)> _
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

Partial Public Class Product

    Private upcField As String
    Private productRxField As ProductRX

    Public Sub New()
        MyBase.New()
        Me.productRxField = New ProductRX
    End Sub

    Public Property upc() As String
        Get
            Return Me.upcField
        End Get
        Set(ByVal value As String)
            Me.upcField = value
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

<DataContract()> _
Partial Public Class ProductRX

    Private sER_IDField As String

    Private pRF_BASECURVEField As Decimal

    Private pRF_DIAMETERField As Decimal

    Private pRD_POWERField As Decimal

    Private pRD_CYLINDERField As Decimal

    Private pRD_AXISField As UShort

    Private pRD_ADDITIONField As String

    Private pRD_COLORField As String

    <DataMember(IsRequired:=True)> _
    Public Property SER_ID() As String
        Get
            Return Me.sER_IDField
        End Get
        Set(ByVal value As String)
            Me.sER_IDField = value
        End Set
    End Property

    <DataMember(IsRequired:=True)> _
    Public Property PRF_BASECURVE() As Decimal
        Get
            Return Me.pRF_BASECURVEField
        End Get
        Set(ByVal value As Decimal)
            Me.pRF_BASECURVEField = value
        End Set
    End Property

    <DataMember(IsRequired:=True)> _
    Public Property PRF_DIAMETER() As Decimal
        Get
            Return Me.pRF_DIAMETERField
        End Get
        Set(ByVal value As Decimal)
            Me.pRF_DIAMETERField = value
        End Set
    End Property

    <DataMember(IsRequired:=True)> _
    Public Property PRD_POWER() As Decimal
        Get
            Return Me.pRD_POWERField
        End Get
        Set(ByVal value As Decimal)
            Me.pRD_POWERField = value
        End Set
    End Property

    <DataMember(IsRequired:=False)> _
    Public Property PRD_CYLINDER() As Decimal
        Get
            Return Me.pRD_CYLINDERField
        End Get
        Set(ByVal value As Decimal)
            Me.pRD_CYLINDERField = value
        End Set
    End Property

    <DataMember(IsRequired:=False)> _
    Public Property PRD_AXIS() As UShort
        Get
            Return Me.pRD_AXISField
        End Get
        Set(ByVal value As UShort)
            Me.pRD_AXISField = value
        End Set
    End Property

    <DataMember(IsRequired:=False)> _
    Public Property PRD_ADDITION() As String
        Get
            Return Me.pRD_ADDITIONField
        End Get
        Set(ByVal value As String)
            Me.pRD_ADDITIONField = value
        End Set
    End Property

    <DataMember(IsRequired:=False)> _
    Public Property PRD_COLOR() As String
        Get
            Return Me.pRD_COLORField
        End Get
        Set(ByVal value As String)
            Me.pRD_COLORField = value
        End Set
    End Property
End Class

#End Region

#Region "OrderResultClasses"

<DataContract()> _
<XmlRoot()> _
Partial Public Class OrderResponse

    Private orderResultsField As List(Of OrderResult)
    Private errorsOccurredField As Boolean
    Private requestErrorsField As List(Of String)

    <DataMember(IsRequired:=True)> _
    <XmlElement()> _
    Public Property errorsOccurred() As Boolean
        Get
            Return Me.errorsOccurredField
        End Get
        Set(ByVal value As Boolean)
            Me.errorsOccurredField = value
        End Set
    End Property

    <DataMember(IsRequired:=False)> _
    <XmlElement(ElementName:="OrderResults")> _
    Public Property OrderResults() As List(Of OrderResult)
        Get
            Return Me.orderResultsField
        End Get
        Set(ByVal value As List(Of OrderResult))
            Me.orderResultsField = value
        End Set
    End Property

    <DataMember(IsRequired:=False)> _
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

<DataContract()> _
Partial Public Class OrderResult
    Private orderIDField As String
    Private orderHasErrorsField As Boolean

    Private orderErrorTextField As List(Of String)

    <DataMember(IsRequired:=True)> _
    <XmlElement()> _
    Public Property orderID() As String
        Get
            Return Me.orderIDField
        End Get
        Set(ByVal value As String)
            Me.orderIDField = value
        End Set
    End Property

    <DataMember(IsRequired:=True)> _
    <XmlElement()> _
    Public Property orderHasErrors() As Boolean
        Get
            Return Me.orderHasErrorsField
        End Get
        Set(ByVal value As Boolean)
            Me.orderHasErrorsField = value
        End Set
    End Property

    <DataMember(IsRequired:=True)> _
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
        Me.orderHasErrors = True
        Me.orderErrors.Add(errorText)
    End Sub

End Class

#End Region
